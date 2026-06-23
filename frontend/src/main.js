// frontend/src/main.js
// Main entry point — imports modules and wires up all event handlers.

import { api }                  from "./utils/api.js";
import { typeIn, copyText, showToast, escHtml } from "./utils/ui.js";
import { seedEvents, todayKey } from "./utils/calendar.js";
import { renderOvMeetings, renderOvUpcoming, updateOverviewStats, initDocChart, updateDocChart, initHealthChart, filterHealthChart } from "./pages/overview.js";
import { renderCalendar, renderSchedule, calNav as _calNav, addEvent as _addEvent, checkAndShowReminder } from "./pages/calendar.js";
import { DOCS, renderDocs, filterDocs, addNewDoc } from "./pages/documents.js";

// ── State ──────────────────────────────────────────────────────
let events = seedEvents();
let healthChart = null;
let _tokenTimer      = null;   // proactive expiry warning timer
let _tokenRefreshing = false;  // prevent concurrent refresh calls

// ── Open Microsoft Teams (desktop app → web fallback) ──────────
window.openTeamsApp = function () {
  // Try launching desktop app via protocol link
  const a = document.createElement("a");
  a.href = "msteams://";
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  // After 2s, if desktop app didn't open, open web version in new tab
  setTimeout(() => window.open("https://teams.cloud.microsoft/", "_blank"), 2000);
};

// ── Navigate to a panel by key ─────────────────────────────────
function navTo(panelKey) {
  const navEl = document.querySelector(`[data-panel="${panelKey}"]`);
  if (navEl) window.nav(navEl);
}

// ── AI particle-network background for the login panel ─────────
let _tmAnim = null;
function startTmBackground() {
  const canvas = document.getElementById("tm-bg-canvas");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");

  function resize() {
    canvas.width  = canvas.offsetWidth  || window.innerWidth;
    canvas.height = canvas.offsetHeight || window.innerHeight;
  }
  resize();
  window.addEventListener("resize", resize);

  const N = 55;
  const pts = Array.from({ length: N }, () => ({
    x:  Math.random() * canvas.width,
    y:  Math.random() * canvas.height,
    vx: (Math.random() - 0.5) * 0.35,
    vy: (Math.random() - 0.5) * 0.35,
    r:  Math.random() * 2.2 + 1,
  }));

  function draw() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    // Lines between close particles
    for (let i = 0; i < N; i++) {
      for (let j = i + 1; j < N; j++) {
        const dx   = pts[i].x - pts[j].x;
        const dy   = pts[i].y - pts[j].y;
        const dist = Math.sqrt(dx * dx + dy * dy);
        if (dist < 160) {
          const alpha = (1 - dist / 160) * 0.22;
          ctx.strokeStyle = `rgba(147,197,253,${alpha})`;
          ctx.lineWidth   = 0.9;
          ctx.beginPath();
          ctx.moveTo(pts[i].x, pts[i].y);
          ctx.lineTo(pts[j].x, pts[j].y);
          ctx.stroke();
        }
      }
    }

    // Glowing dots
    pts.forEach(p => {
      const g = ctx.createRadialGradient(p.x, p.y, 0, p.x, p.y, p.r * 4);
      g.addColorStop(0,   "rgba(191,219,254,0.95)");
      g.addColorStop(0.4, "rgba(99,179,237,0.4)");
      g.addColorStop(1,   "rgba(99,179,237,0)");
      ctx.beginPath();
      ctx.arc(p.x, p.y, p.r * 4, 0, Math.PI * 2);
      ctx.fillStyle = g;
      ctx.fill();

      p.x += p.vx;  p.y += p.vy;
      if (p.x < 0 || p.x > canvas.width)  p.vx *= -1;
      if (p.y < 0 || p.y > canvas.height) p.vy *= -1;
    });

    _tmAnim = requestAnimationFrame(draw);
  }

  if (_tmAnim) cancelAnimationFrame(_tmAnim);
  _tmAnim = requestAnimationFrame(draw);
}

function stopTmBackground() {
  if (_tmAnim) { cancelAnimationFrame(_tmAnim); _tmAnim = null; }
}

// ── Teams Chat ─────────────────────────────────────────────────
window.tcToggleTokenInput = function () {
  const area = document.getElementById("tc-token-input-area");
  if (area) area.style.display = area.style.display === "none" ? "block" : "none";
};

window.tcSaveToken = function () {
  const val = document.getElementById("tc-token-input")?.value?.trim();
  if (!val) { showToast("Paste your token first."); return; }
  localStorage.setItem("spToken", val);
  updateTokenUI();
  const area = document.getElementById("tc-token-input-area");
  if (area) area.style.display = "none";
  showToast("Microsoft 365 connected.");
  window.syncTeamsChats(document.getElementById("tc-sync-btn"));
};

window.syncTeamsChats = async function (btn) {
  const token     = localStorage.getItem("spToken") || "";
  const noTokEl   = document.getElementById("tc-no-token");
  const chatUiEl  = document.getElementById("tc-chat-ui");

  if (!token) {
    if (noTokEl)  noTokEl.style.display  = "block";
    if (chatUiEl) chatUiEl.style.display = "none";
    showToast("Paste a Microsoft 365 token to sync Teams chats.");
    return;
  }

  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader" style="font-size:12px;animation:spin .7s linear infinite"></i> Syncing…'; }

  const data = await api.teamsChats();

  if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-refresh"></i> Sync chats'; }

  if (data.error) {
    if (data.noPermission) {
      // Chat.Read scope was missing — show permission banner instead of "no token" card
      if (noTokEl) {
        noTokEl.style.display = "block";
        const msgEl = noTokEl.querySelector(".tc-no-token-msg");
        if (msgEl) msgEl.textContent = "Chat.Read permission is missing. Sign out and reconnect via Microsoft 365 to get a fresh token with Teams chat access. If it persists, ask your IT admin to grant Chat.Read for this app.";
      }
      if (chatUiEl) chatUiEl.style.display = "none";
      showToast("Teams: Chat.Read permission not granted — reconnect via Microsoft 365.");
    } else {
      if (noTokEl)  noTokEl.style.display  = "block";
      if (chatUiEl) chatUiEl.style.display = "none";
      showToast("Teams chat error: " + data.error);
    }
    return;
  }

  if (noTokEl)  noTokEl.style.display  = "none";
  if (chatUiEl) chatUiEl.style.display = "block";

  const chats   = data.chats || [];
  const syncLbl = document.getElementById("tc-sync-lbl");
  if (syncLbl) syncLbl.textContent = "Last synced: " + new Date().toLocaleTimeString("en-IN");
  const countEl = document.getElementById("tc-chat-count");
  if (countEl) { countEl.textContent = chats.length; countEl.style.display = chats.length ? "inline" : "none"; }

  _renderTcChatList(chats);
  showToast(`Loaded ${chats.length} conversation(s) from Teams.`);
};

let _tcChats = [];
function _renderTcChatList(chats) {
  _tcChats = chats;
  const listEl = document.getElementById("tc-chat-list");
  if (!listEl) return;

  if (!chats.length) {
    listEl.innerHTML = '<div style="font-size:12px;color:var(--hint);padding:20px 0;text-align:center">No conversations found.</div>';
    return;
  }

  listEl.innerHTML = "";
  chats.forEach((chat, i) => {
    const memberNames = (chat.members || []).map(m => m.displayName).filter(Boolean);
    const name   = chat.topic || memberNames.slice(0, 2).join(", ") || "Teams chat";
    const preview = chat.lastMessagePreview?.body?.content?.replace(/<[^>]+>/g, "").trim() || "";
    const time   = chat.lastUpdatedDateTime
      ? new Date(chat.lastUpdatedDateTime).toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" })
      : "";
    const initials = name.split(" ").slice(0, 2).map(w => w[0] || "").join("").toUpperCase() || "T";

    const row = document.createElement("div");
    row.className = "meet-row";
    row.style.gap = "10px";
    row.innerHTML = `
      <div style="width:36px;height:36px;background:#6264a7;border-radius:9px;display:flex;align-items:center;justify-content:center;flex-shrink:0;font-size:13px;font-weight:700;color:#fff">${escHtml(initials)}</div>
      <div style="flex:1;min-width:0">
        <div class="mr-t">${escHtml(name)}</div>
        <div class="mr-m" style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escHtml(preview.slice(0, 55))}${preview.length > 55 ? "…" : ""}</div>
      </div>
      ${time ? `<span style="font-size:10px;color:var(--hint);flex-shrink:0">${escHtml(time)}</span>` : ""}`;
    const teamsUrl = `https://teams.microsoft.com/l/chat/${encodeURIComponent(chat.id)}/0`;
    row.title = "Open in Microsoft Teams";
    row.style.cursor = "pointer";
    row.onclick = () => window.open(teamsUrl, "_blank");
    listEl.appendChild(row);
  });
}

async function _loadTcMessages(chatId, chatName, row) {
  document.querySelectorAll("#tc-chat-list .meet-row").forEach(r => r.classList.remove("sel"));
  if (row) row.classList.add("sel");

  const panel = document.getElementById("tc-messages-panel");
  if (!panel) return;

  panel.style.display    = "block";
  panel.style.alignItems = "initial";
  panel.style.justifyContent = "initial";
  panel.innerHTML = `
    <div style="font-size:13px;font-weight:600;color:var(--text);margin-bottom:10px;padding-bottom:9px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:8px">
      <i class="ti ti-brand-teams" style="color:#6264a7;font-size:15px"></i> ${escHtml(chatName)}
    </div>
    <div style="font-size:12px;color:var(--muted);text-align:center;padding:20px 0">
      <span class="ldot" style="background:#6264a7"></span> Loading messages…
    </div>`;

  const data = await api.teamsChatMessages(chatId);

  if (data.error) {
    panel.innerHTML = `<div style="font-size:12px;color:var(--red);padding:12px;background:var(--red-bg);border-radius:var(--r-sm)">${escHtml(data.error)}</div>`;
    return;
  }

  const messages = (data.messages || []).filter(m => m.messageType !== "systemEventMessage").reverse();

  const teamsUrl = `https://teams.microsoft.com/l/chat/${encodeURIComponent(chatId)}/0`;
  panel.innerHTML = `
    <div style="font-size:13px;font-weight:600;color:var(--text);margin-bottom:10px;padding-bottom:9px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:8px">
      <i class="ti ti-brand-teams" style="color:#6264a7;font-size:15px"></i> ${escHtml(chatName)}
      <span class="badge" style="background:rgba(98,100,167,.12);color:#6264a7;border-color:rgba(98,100,167,.25)">${messages.length} messages</span>
      <a href="${escHtml(teamsUrl)}" target="_blank" style="margin-left:auto;text-decoration:none"><button class="sm" style="font-size:11px;background:#6264a7;color:#fff;border-color:#6264a7;display:inline-flex;align-items:center;gap:4px"><i class="ti ti-external-link" style="font-size:11px"></i> Open in Teams</button></a>
    </div>
    <div id="tc-msg-list" style="display:flex;flex-direction:column;gap:12px;max-height:500px;overflow-y:auto;padding-right:4px">
      ${!messages.length ? '<div style="font-size:12px;color:var(--hint);text-align:center;padding:16px">No messages in this conversation.</div>' : ""}
    </div>`;

  const listEl = panel.querySelector("#tc-msg-list");
  messages.forEach(msg => {
    const sender   = msg.from?.user?.displayName || msg.from?.application?.displayName || "Unknown";
    const body     = msg.body?.content?.replace(/<[^>]+>/g, "").trim() || "(no content)";
    const time     = msg.createdDateTime
      ? new Date(msg.createdDateTime).toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" }) : "";
    const date     = msg.createdDateTime
      ? new Date(msg.createdDateTime).toLocaleDateString("en-IN", { day: "2-digit", month: "short" }) : "";
    const initials = sender.split(" ").slice(0, 2).map(w => w[0] || "").join("").toUpperCase() || "?";

    const el = document.createElement("div");
    el.style.cssText = "display:flex;gap:9px;align-items:flex-start";
    el.innerHTML = `
      <div style="width:30px;height:30px;background:#6264a7;border-radius:7px;display:flex;align-items:center;justify-content:center;flex-shrink:0;font-size:11px;font-weight:700;color:#fff;margin-top:1px">${escHtml(initials)}</div>
      <div style="flex:1;min-width:0">
        <div style="display:flex;align-items:baseline;gap:7px;margin-bottom:3px">
          <span style="font-size:12px;font-weight:600;color:var(--text)">${escHtml(sender)}</span>
          <span style="font-size:10px;color:var(--hint)">${escHtml(date)} ${escHtml(time)}</span>
        </div>
        <div style="font-size:12px;color:var(--muted);line-height:1.55;word-break:break-word">${escHtml(body.slice(0, 500))}${body.length > 500 ? "…" : ""}</div>
      </div>`;
    listEl.appendChild(el);
  });

  if (listEl) listEl.scrollTop = listEl.scrollHeight;
}

// ── Navigation ─────────────────────────────────────────────────
const PAGE_TITLES = {
  tm: "Teams meetings",
  ov: "Dashboard",        cal: "Calendar & meetings",
  tc: "Teams Chat",       ol:  "Outlook inbox",
  wr: "Weekly report",    mm:  "MOM generator",
  su: "Stand-up generator", dc: "Document repository",
  lv: "Leave & resources",  ai: "General assistant",
};

function isProjectConfigured() {
  return !!(localStorage.getItem("projectName") || "").trim();
}
function applyNavLock() {
  const ok = isProjectConfigured();
  document.querySelectorAll(".nav-item[data-panel]").forEach(el => {
    if (el.dataset.panel === "ov") return;
    el.classList.toggle("nav-locked", !ok);
  });
}

window.nav = function (el) {
  const id = el.dataset.panel;
  if (id !== "ov" && !isProjectConfigured()) return;
  document.querySelectorAll(".panel").forEach((p) => p.classList.remove("active"));
  document.querySelectorAll(".nav-item").forEach((n) => n.classList.remove("active"));
  const panel   = document.getElementById("p-" + id);
  const content = document.getElementById("content");
  if (panel) panel.classList.add("active");
  el.classList.add("active");
  const titleEl = document.getElementById("page-title");
  if (titleEl) titleEl.textContent = PAGE_TITLES[id] || id;
  const appEl = document.getElementById("app");
  const isLogin = id === "tm" && !ewsGetCreds().ewsUrl;
  if (appEl) appEl.classList.toggle("login-mode", isLogin);
  if (isLogin) { startTmBackground(); } else { stopTmBackground(); }

  if (id === "tc") {
    const tok      = localStorage.getItem("spToken") || "";
    const noTokEl  = document.getElementById("tc-no-token");
    const chatUiEl = document.getElementById("tc-chat-ui");
    if (tok) {
      if (noTokEl) noTokEl.style.display = "none";
      if (chatUiEl && chatUiEl.style.display === "none" && document.getElementById("tc-chat-list")?.children.length <= 1) {
        window.syncTeamsChats(document.getElementById("tc-sync-btn"));
      } else if (chatUiEl) {
        chatUiEl.style.display = "block";
      }
    } else {
      if (noTokEl)  noTokEl.style.display  = "block";
      if (chatUiEl) chatUiEl.style.display = "none";
    }
  }
  if (id === "cal") {
    renderCalendar(events);
    renderSchedule(todayKey(), events);
    const creds = ewsGetCreds();
    if (creds.ewsUrl && !_ewsMeetings.length) {
      window.syncOutlookCalendar(document.getElementById("cal-sync-btn"));
    }
  }

  if (id === "dc") {
    window.syncSharePointDocs(document.getElementById("sp-sync-btn"));
  }
};

// ── Clock ──────────────────────────────────────────────────────
function getISOWeek(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  date.setUTCDate(date.getUTCDate() + 4 - (date.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  return { week: Math.ceil(((date - yearStart) / 86400000 + 1) / 7), year: date.getUTCFullYear() };
}

function tick() {
  const n = new Date();
  const clockEl   = document.getElementById("clk");
  const dateEl    = document.getElementById("dt-lbl");
  const tbTime    = document.getElementById("topbar-time");
  const tbWeekEl  = document.getElementById("topbar-week");
  if (clockEl)  clockEl.textContent  = n.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
  if (dateEl)   dateEl.textContent   = n.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
  if (tbTime)   tbTime.textContent   = n.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" });
  if (tbWeekEl) { const { week, year } = getISOWeek(n); tbWeekEl.textContent = `Wk ${week} · ${year}`; }
  checkAndShowReminder(events);
}
setInterval(tick, 1000);
tick();

// ── Calendar ───────────────────────────────────────────────────
window.calNav      = (dir) => { _calNav(dir, events); };
window.addEvent = async () => {
  // 1 — add locally (existing behaviour)
  events = _addEvent(events);
  renderCalendar(events);
  renderSchedule(todayKey(), events);
  renderOvMeetings(events);
  document.getElementById("cal-badge").textContent = (events[todayKey()] || []).length;

  // 2 — send Outlook invite via Exchange if connected
  const creds = ewsGetCreds();
  const statusEl = document.getElementById("ev-invite-status");
  const btn      = document.getElementById("ev-add-btn");

  if (!creds.ewsUrl) {
    if (statusEl) { statusEl.style.display = "block"; statusEl.style.background = "rgba(220,150,0,.1)"; statusEl.style.color = "var(--orange,#d97706)"; statusEl.textContent = "⚠ Not connected to Exchange — event saved locally only. Connect in the Teams panel to send invites."; }
    setTimeout(() => { if (statusEl) statusEl.style.display = "none"; }, 6000);
    return;
  }

  const title    = document.getElementById("ev-title")?.value?.trim();
  const date     = document.getElementById("ev-date")?.value || new Date().toISOString().slice(0, 10);
  const time     = document.getElementById("ev-time")?.value?.trim();
  const dur      = document.getElementById("ev-dur")?.value;
  const attRaw   = document.getElementById("ev-att")?.value?.trim();
  const location = document.getElementById("ev-location")?.value?.trim();
  const desc     = document.getElementById("ev-desc")?.value?.trim();

  if (!title) {
    if (statusEl) { statusEl.style.display = "block"; statusEl.style.background = "rgba(220,38,38,.07)"; statusEl.style.color = "var(--red)"; statusEl.textContent = "✗ Please enter a meeting subject."; }
    setTimeout(() => { if (statusEl) statusEl.style.display = "none"; }, 4000);
    return;
  }
  const attendees = attRaw ? attRaw.split(/[,;]+/).map(a => a.trim()).filter(a => a.includes("@")) : [];

  if (statusEl) { statusEl.style.display = "block"; statusEl.style.background = "var(--surface2)"; statusEl.style.color = "var(--muted)"; statusEl.textContent = "Sending invite via Exchange…"; }
  if (btn) btn.disabled = true;

  const data = await api.ewsCreateMeeting({
    ...creds,
    subject:   title,
    date,
    time,
    duration:  dur,
    attendees,
    location,
    body:      desc,
  });

  if (btn) btn.disabled = false;
  if (data.ok) {
    const sent = (data.sentTo && data.sentTo.length) ? data.sentTo : attendees;
    const recipientMsg = sent.length ? `Invite sent to: ${sent.join(", ")}` : "Saved to your calendar";
    if (statusEl) { statusEl.style.display = "block"; statusEl.style.background = "var(--green-bg)"; statusEl.style.color = "var(--green)"; statusEl.textContent = `✓ ${recipientMsg}`; }
    showToast(`Meeting "${title}" created & invite sent via Exchange!`);
    const descEl = document.getElementById("ev-desc");    if (descEl) descEl.value = "";
    const locEl  = document.getElementById("ev-location"); if (locEl) locEl.value = "";
  } else {
    if (statusEl) { statusEl.style.display = "block"; statusEl.style.background = "rgba(220,38,38,.07)"; statusEl.style.color = "var(--red)"; statusEl.textContent = `✗ ${data.error}`; }
  }
  setTimeout(() => { if (statusEl) statusEl.style.display = "none"; }, 6000);
};

// ── Overview charts ────────────────────────────────────────────
window.updateHealthChart = () => {
  const v = document.getElementById("ph-filter")?.value;
  if (healthChart) filterHealthChart(healthChart, v);
};

// ── Document repo ──────────────────────────────────────────────
window.filterDocs = filterDocs;
window.addNewDoc  = addNewDoc;

// Show chosen filename in the upload form label
window.onDocFileChosen = function (input) {
  const lbl = document.getElementById("dc-file-name");
  if (lbl) lbl.textContent = input.files[0]?.name || "Choose file";
  // Pre-fill path from last used if empty
  const pathEl = document.getElementById("dc-upload-path");
  if (pathEl && !pathEl.value) pathEl.value = localStorage.getItem("spDocFolder") || "Documents/BA Hub";
};

// Upload from the inline form (file input + path field)
window.uploadDocFromForm = async function (btn) {
  const fileInput = document.getElementById("dc-file-input");
  const pathEl    = document.getElementById("dc-upload-path");
  const statusEl  = document.getElementById("dc-upload-status");
  const file      = fileInput?.files[0];
  const path      = pathEl?.value?.trim() || "Documents/BA Hub";

  if (!file) { showToast("Choose a file first."); return; }

  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader" style="font-size:12px;animation:spin .7s linear infinite"></i>'; }
  if (statusEl) { statusEl.style.display = "block"; statusEl.textContent = "Uploading…"; }

  localStorage.setItem("spDocFolder", path);

  const today = new Date().toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
  const name  = file.name.replace(/\.[^.]+$/, "");
  DOCS.unshift({ n: name, v: "v1.0", s: "Pending", d: today, desc: "Uploaded from local file.", url: "" });
  filterDocs();

  const token = localStorage.getItem("spToken") || "";
  if (!token) {
    if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-upload"></i> Upload'; }
    if (statusEl) {
      statusEl.style.display = "block";
      statusEl.innerHTML = `<i class="ti ti-alert-circle" style="color:var(--orange);margin-right:4px"></i>Microsoft 365 not connected — <a href="/api/auth/microsoft" style="color:#0078d4;font-weight:600;text-decoration:none">Connect now</a> to upload to SharePoint.`;
    }
    return;
  }

  const fd = new FormData();
  fd.append("file",       file);
  fd.append("token",      token);
  fd.append("folderPath", path);
  const data = await api.uploadDoc(fd);

  if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-upload"></i> Upload'; }

  if (data.error) {
    if (statusEl) { statusEl.style.display = "block"; statusEl.textContent = "Upload failed: " + data.error; }
    showToast("Upload failed: " + data.error, 4000);
    return;
  }

  DOCS[0].url = data.url || "";
  DOCS[0].s   = "Approved";
  try { localStorage.setItem("ba_docs", JSON.stringify(DOCS)); } catch {}
  filterDocs();

  // Reset form
  fileInput.value = "";
  const lbl = document.getElementById("dc-file-name");
  if (lbl) lbl.textContent = "Choose file (.pdf / .docx)";
  if (statusEl) { statusEl.style.display = "block"; statusEl.textContent = `Uploaded: ${data.name} → ${path}`; }
  showToast(`Uploaded: ${data.name}`);
};

// ── SharePoint source toggle (recent vs by-site) ──────────────────
let _spSource = "recent";

window.spSetSource = function (src, btn) {
  _spSource = src;
  const recentBtn = document.getElementById("sp-src-recent");
  const siteBtn   = document.getElementById("sp-src-site");
  const siteRow   = document.getElementById("sp-site-row");
  if (recentBtn) { recentBtn.style.background = src === "recent" ? "var(--blue)" : "var(--surface2)"; recentBtn.style.color = src === "recent" ? "#fff" : "var(--muted)"; }
  if (siteBtn)   { siteBtn.style.background   = src === "site"   ? "var(--blue)" : "var(--surface2)"; siteBtn.style.color   = src === "site"   ? "#fff" : "var(--muted)"; }
  if (siteRow)   { siteRow.style.display = src === "site" ? "flex" : "none"; }
  if (src === "site") {
    const sel = document.getElementById("sp-site-select");
    if (sel && sel.options.length <= 1) window.loadSpSites(null);
  }
};

window.loadSpSites = async function (btn) {
  const sel      = document.getElementById("sp-site-select");
  const statusEl = document.getElementById("sp-sync-status");
  if (!localStorage.getItem("spToken")) {
    if (statusEl) {
      statusEl.style.display = "block";
      statusEl.style.background = "var(--surface2)";
      statusEl.style.color = "var(--text)";
      statusEl.innerHTML =
        '<div style="margin-bottom:8px;font-size:13px;font-weight:500"><i class="ti ti-brand-office" style="color:#0078d4"></i> Connect Microsoft 365 to load SharePoint sites</div>' +
        '<div style="font-size:12px;color:var(--muted);margin-bottom:8px;line-height:1.6">Go to <a href="https://developer.microsoft.com/en-us/graph/graph-explorer" target="_blank" style="color:var(--blue)">Graph Explorer</a>, sign in with your Microsoft 365 account, then copy the <strong>Access token</strong> from the Auth tab.</div>' +
        '<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">' +
        '<input id="sp-inline-token" type="password" placeholder="Paste Microsoft 365 access token…" style="flex:1;min-width:200px;margin:0;font-size:12px;font-family:var(--mono)">' +
        '<button class="sm primary" onclick="spConnectInline()"><i class="ti ti-plug"></i> Connect &amp; Load</button>' +
        '</div>';
    }
    return;
  }
  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader" style="font-size:12px;animation:spin .7s linear infinite"></i>'; }
  if (statusEl) { statusEl.style.display = "block"; statusEl.style.background = "var(--surface2)"; statusEl.style.color = "var(--muted)"; statusEl.textContent = "Loading SharePoint sites…"; }
  const data = await api.spSites();
  if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-refresh"></i> Load sites'; }
  if (data.error) {
    if (statusEl) { statusEl.style.background = "rgba(220,38,38,.07)"; statusEl.style.color = "var(--red)"; statusEl.textContent = "Sites error: " + data.error; }
    showToast("Sites error: " + data.error);
    return;
  }
  if (!sel) return;
  sel.innerHTML = '<option value="">— select a site —</option>';
  (data.sites || []).forEach(function (s) {
    const opt = document.createElement("option");
    opt.value = s.id; opt.textContent = s.name; sel.appendChild(opt);
  });
  const src = data.source === "root" ? " (root only — add Sites.Read.All for full list)" : data.source === "followed" ? " (followed sites)" : "";
  const msg = "Loaded " + (data.sites || []).length + " site(s)" + src;
  if (statusEl) { statusEl.style.background = "var(--green-bg)"; statusEl.style.color = "var(--green-text)"; statusEl.textContent = msg; }
  showToast(msg);
};

window.spOnSiteChange = function () {
  const siteId = document.getElementById("sp-site-select")?.value;
  if (siteId) window.syncSharePointDocs(document.getElementById("sp-sync-btn"));
};

// ── Connect M365 token inline from Documents panel ────────────────
window.spConnectInline = function () {
  const inp = document.getElementById("sp-inline-token");
  const val = inp ? inp.value.trim() : "";
  if (!val) { showToast("Paste your Microsoft 365 token first."); return; }
  localStorage.setItem("spToken", val);
  updateTokenUI();
  if (typeof scheduleTokenRefreshTimer === "function") scheduleTokenRefreshTimer(val);
  showToast("Microsoft 365 connected.");
  const statusEl = document.getElementById("sp-sync-status");
  if (statusEl) statusEl.innerHTML = "";
  // If in site mode, load sites; otherwise sync recent files
  if (_spSource === "site") {
    window.loadSpSites(document.getElementById("sp-sync-btn"));
  } else {
    window.syncSharePointDocs(document.getElementById("sp-sync-btn"));
  }
};

// ── Sync documents from SharePoint (Microsoft Graph API) ──────────
window.syncSharePointDocs = async function (btn) {
  const token = localStorage.getItem("spToken") || "";
  const statusEl = document.getElementById("sp-sync-status");
  const userEl   = document.getElementById("sp-sync-user");

  if (!token) {
    if (statusEl) {
      statusEl.style.display = "block";
      statusEl.style.background = "var(--surface2)";
      statusEl.style.color = "var(--text)";
      statusEl.innerHTML =
        '<div style="margin-bottom:10px;font-size:13px;font-weight:500"><i class="ti ti-brand-office" style="color:#0078d4"></i> Connect Microsoft 365 to fetch your SharePoint documents</div>' +
        '<div style="font-size:12px;color:var(--muted);margin-bottom:10px;line-height:1.6">Get a token from <a href="https://developer.microsoft.com/en-us/graph/graph-explorer" target="_blank" style="color:var(--blue)">Graph Explorer</a> → sign in → copy the <strong>Access token</strong> tab.</div>' +
        '<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">' +
        '<input id="sp-inline-token" type="password" placeholder="Paste Microsoft 365 access token here…" style="flex:1;min-width:200px;margin:0;font-size:12px;font-family:var(--mono)">' +
        '<button class="sm primary" onclick="spConnectInline()"><i class="ti ti-plug"></i> Connect</button>' +
        '</div>';
    }
    return;
  }

  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader" style="font-size:12px;animation:spin .7s linear infinite"></i> Syncing…'; }
  if (statusEl) { statusEl.style.display = "block"; statusEl.style.background = "var(--surface2)"; statusEl.style.color = "var(--muted)"; statusEl.textContent = "Fetching files from SharePoint…"; }

  // Branch: recent files (OneDrive) OR site document libraries
  let data;
  if (_spSource === "site") {
    const siteId = document.getElementById("sp-site-select")?.value;
    if (!siteId) {
      if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-refresh"></i> Sync'; }
      if (statusEl) { statusEl.textContent = "Select a site first."; }
      return;
    }
    if (statusEl) statusEl.textContent = "Fetching document libraries from site…";
    data = await api.spSiteFiles(siteId);
  } else {
    data = await api.spFiles();
  }

  if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-refresh"></i> Sync'; }

  if (data.error) {
    if (statusEl) { statusEl.style.background = "rgba(220,38,38,.07)"; statusEl.style.color = "var(--red)"; statusEl.textContent = data.error; }
    return;
  }

  const spFiles = data.files || [];
  const today = new Date().toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
  const existingSpIds = new Set(DOCS.filter(function(d) { return d._spId; }).map(function(d) { return d._spId; }));
  let added = 0, updated = 0;

  spFiles.forEach(function(f) {
    const modified = f.modified ? new Date(f.modified).toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" }) : today;
    const extMatch = f.name.match(/\.(\w+)$/);
    const ext = extMatch ? extMatch[1].toLowerCase() : "";
    const status = (ext === "pdf" || ext === "docx" || ext === "doc") ? "Approved" : "Pending";
    if (existingSpIds.has(f.id)) {
      const idx = DOCS.findIndex(function(d) { return d._spId === f.id; });
      if (idx >= 0) { DOCS[idx].d = modified; DOCS[idx].url = f.webUrl; updated++; }
    } else {
      DOCS.unshift({ n: f.name, v: "v1.0", s: status, d: modified, desc: "SharePoint — " + f.folder, url: f.webUrl, _spId: f.id });
      added++;
    }
  });

  try { localStorage.setItem("ba_docs", JSON.stringify(DOCS)); } catch {}
  filterDocs();

  const srcLabel = _spSource === "site" ? ("site (" + (data.libCount || 0) + " libraries)") : "OneDrive recent";
  const msg = "Synced " + spFiles.length + " file(s) from " + srcLabel + " — " + added + " new, " + updated + " updated.";
  if (statusEl) { statusEl.style.background = "var(--green-bg)"; statusEl.style.color = "var(--green-text)"; statusEl.textContent = msg; }
  if (userEl)   { const uname = localStorage.getItem("ewsUsername") || ""; if (uname) userEl.textContent = uname; }
  showToast(msg);
};

// Upload via SharePoint (legacy — called from documents.js row menu)
window._uploadDocFile = async function (idx, file, folderPath) {
  const token = localStorage.getItem("spToken") || "";
  if (!token) {
    // No token — still add the file locally and prompt to set SharePoint URL manually
    if (idx >= 0 && idx < DOCS.length) {
      const today = new Date().toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
      DOCS[idx].s = "Pending";
      DOCS[idx].d = today;
    }
    filterDocs();
    showToast("File saved locally. Paste the SharePoint link in the document's Source link field, or set up OAuth in Weekly Report → Settings.");
    return;
  }
  showToast("Uploading to SharePoint...");
  const fd = new FormData();
  fd.append("file",       file);
  fd.append("token",      token);
  fd.append("folderPath", folderPath);
  const data = await api.uploadDoc(fd);
  if (data.error) {
    showToast("Upload failed: " + data.error + " — paste the SharePoint URL manually in the Source link field.");
    return;
  }
  if (idx >= 0 && idx < DOCS.length) {
    DOCS[idx].url = data.url || "";
    DOCS[idx].s   = "Approved";
    const today = new Date().toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
    DOCS[idx].d   = today;
  }
  filterDocs();
  showToast(`Uploaded: ${data.name} → SharePoint`);
};

window.triggerNewDocUpload = function () {
  const input = document.createElement("input");
  input.type   = "file";
  input.accept = ".pdf,.doc,.docx";
  input.onchange = () => {
    if (!input.files[0]) return;
    let path = localStorage.getItem("spDocFolder");
    if (!path) {
      path = prompt("SharePoint / OneDrive folder path:", "Documents/Falcon Dashboard");
      if (!path) return;
      localStorage.setItem("spDocFolder", path);
    }
    const today = new Date().toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
    const name  = input.files[0].name.replace(/\.[^.]+$/, "");
    DOCS.unshift({ n: name, v: "v1.0", s: "Pending", d: today, desc: "Uploaded file.", url: "" });
    filterDocs();
    window._showDocDetail(0);  // open detail panel so source link field is visible
    window._uploadDocFile(0, input.files[0], path);
  };
  input.click();
};

// ── Microsoft 365 token (Graph Explorer, no admin needed) ─────
function getMsToken() {
  return localStorage.getItem("spToken") || "";
}
function updateTokenUI() {
  const tok         = getMsToken();
  const badge       = document.getElementById("tm-connected-badge");
  const hint        = document.getElementById("ol-token-hint");
  const ok          = document.getElementById("ol-token-ok");
  const inp         = document.getElementById("ms-token-input");
  const dcHint      = document.getElementById("dc-connect-hint");
  if (badge)  badge.style.display  = tok ? "inline-flex" : "none";
  if (hint)   hint.style.display   = tok ? "none"        : "inline";
  if (ok)     ok.style.display     = tok ? "inline"      : "none";
  if (inp && !inp.value && tok)  inp.value = tok;
  if (dcHint) dcHint.style.display = tok ? "none"        : "block";
}
window.connectMsToken = function () {
  const val = document.getElementById("ms-token-input")?.value?.trim();
  if (!val) { showToast("Paste your token first."); return; }
  // Reject tokens that are already expired
  const expiry = parseJwtExpiry(val);
  if (expiry && expiry <= new Date()) {
    showToast("This token is already expired. Get a fresh token from Graph Explorer and paste again.", 5000);
    return;
  }
  if (expiry) {
    const minLeft = Math.round((expiry - Date.now()) / 60000);
    if (minLeft < 10) showToast(`Token expires in ${minLeft} min — fetch a fresh one soon.`);
  }
  localStorage.setItem("spToken", val);
  scheduleTokenRefreshTimer(val);
  updateTokenUI();
  showToast("Microsoft 365 connected.");
};
window.clearMsToken = function () {
  localStorage.removeItem("spToken");
  localStorage.removeItem("spRefreshToken");
  if (_tokenTimer) { clearTimeout(_tokenTimer); _tokenTimer = null; }
  const inp = document.getElementById("ms-token-input");
  if (inp) inp.value = "";
  updateTokenUI();
  showToast("Token cleared.");
};

// ── Token lifecycle helpers ────────────────────────────────────
function parseJwtExpiry(token) {
  try {
    const b64 = token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
    const payload = JSON.parse(atob(b64));
    return payload.exp ? new Date(payload.exp * 1000) : null;
  } catch { return null; }
}

function scheduleTokenRefreshTimer(token) {
  if (_tokenTimer) clearTimeout(_tokenTimer);
  const expiry = parseJwtExpiry(token);
  if (!expiry) return;
  // Fire 5 min before expiry so we can silently refresh
  const delay = expiry.getTime() - 5 * 60 * 1000 - Date.now();
  if (delay <= 0) { handleTokenExpired(); return; }
  _tokenTimer = setTimeout(handleTokenExpired, delay);
}

async function handleTokenExpired() {
  if (_tokenRefreshing) return;
  _tokenRefreshing = true;
  const rt = localStorage.getItem("spRefreshToken");
  if (rt) {
    try {
      const data = await api.refreshToken(rt);
      if (data.accessToken) {
        localStorage.setItem("spToken", data.accessToken);
        if (data.refreshToken) localStorage.setItem("spRefreshToken", data.refreshToken);
        updateTokenUI();
        scheduleTokenRefreshTimer(data.accessToken);
        _tokenRefreshing = false;
        showToast("Microsoft 365 token refreshed automatically.");
        return;
      }
    } catch (_) { /* fall through */ }
  }
  // Cannot auto-refresh — clear and ask user to reconnect
  _tokenRefreshing = false;
  localStorage.removeItem("spToken");
  localStorage.removeItem("spRefreshToken");
  if (_tokenTimer) { clearTimeout(_tokenTimer); _tokenTimer = null; }
  updateTokenUI();
  showToast("Microsoft 365 session expired — please reconnect in Weekly Report → Settings → OAuth setup.", 6000);
}

// Listen for expired-token events fired by api.js
window.addEventListener("ms-token-expired", () => handleTokenExpired());

// ── ICS Calendar Parser ────────────────────────────────────────
function parseICS(text) {
  const meetings = [];
  const blocks   = text.split("BEGIN:VEVENT");
  blocks.shift(); // remove header block

  for (const block of blocks) {
    const get = (key) => {
      const re  = new RegExp(`${key}(?:;[^:]*)?:([^\\r\\n]+)`, "i");
      const m   = block.match(re);
      return m ? m[1].trim() : "";
    };
    const getMulti = (key) => {
      const re  = new RegExp(`${key}(?:;[^:]*)?:([^\\r\\n]+)`, "gi");
      const out = [];
      let m;
      while ((m = re.exec(block)) !== null) out.push(m[1].trim());
      return out;
    };

    const summary   = get("SUMMARY").replace(/\\,/g, ",").replace(/\\n/g, " ");
    const dtstart   = get("DTSTART");
    const dtend     = get("DTEND");
    const location  = get("LOCATION");
    const desc      = get("DESCRIPTION").replace(/\\n/g, " ").replace(/\\,/g, ",");
    const attendees = getMulti("ATTENDEE")
      .map((a) => {
        const cn   = a.match(/CN=([^:;]+)/i);
        const mail = a.match(/mailto:(.+)/i);
        return cn ? cn[1] : (mail ? mail[1] : a);
      })
      .filter(Boolean);

    if (!summary || !dtstart) continue;

    // Parse datetime (handles YYYYMMDDTHHMMSSZ and YYYYMMDD)
    const parseDate = (s) => {
      if (!s) return null;
      const clean = s.replace(/[TZ]/g, "").replace(/[^0-9]/g, "");
      if (clean.length >= 8) {
        return new Date(
          parseInt(clean.slice(0,4)),
          parseInt(clean.slice(4,6)) - 1,
          parseInt(clean.slice(6,8)),
          parseInt(clean.slice(8,10) || "0"),
          parseInt(clean.slice(10,12) || "0")
        );
      }
      return null;
    };

    const start = parseDate(dtstart);
    const end   = parseDate(dtend);
    if (!start) continue;

    const dur = (start && end)
      ? `${Math.round((end - start) / 60000)} min`
      : "N/A";

    meetings.push({ summary, start, end, dur, location, desc, attendees });
  }

  // Sort by start date, show upcoming first then recent past
  const now = new Date();
  return meetings
    .sort((a, b) => Math.abs(a.start - now) - Math.abs(b.start - now))
    .slice(0, 50);
}

function renderICSMeetings(meetings) {
  const list     = document.getElementById("teams-list");
  const countEl  = document.getElementById("meetings-count");
  if (!list) return;

  if (!meetings.length) {
    list.innerHTML = '<div style="font-size:12px;color:var(--hint);padding:16px 0;text-align:center">No meetings found in this calendar file.</div>';
    return;
  }

  if (countEl) { countEl.style.display = "inline"; countEl.textContent = meetings.length; }

  const now = new Date();
  list.innerHTML = "";
  meetings.forEach((m, i) => {
    const isPast   = m.start < now;
    const dateStr  = m.start.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
    const timeStr  = m.start.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" });
    const isTeams  = (m.location || "").toLowerCase().includes("teams") || (m.desc || "").toLowerCase().includes("teams");
    const row      = document.createElement("div");
    row.className  = "meet-row";
    row.innerHTML  = `
      <div style="flex:1;min-width:0">
        <div class="mr-t">${escHtml(m.summary)}</div>
        <div class="mr-m">${dateStr} · ${timeStr} · ${m.dur}${m.attendees.length ? " · " + escHtml(m.attendees.slice(0,3).join(", ")) : ""}</div>
      </div>
      <span class="badge ${isPast ? "b-green" : "b-blue"}">${isPast ? "Past" : "Upcoming"}</span>
      ${isTeams ? '<span class="badge b-blue" style="margin-left:4px">Teams</span>' : ""}`;
    row.onclick = () => {
      const teamsUrl = (m.desc || "").match(/https:\/\/teams\.microsoft\.com\/[^\s<>"]+/)?.[0];
      if (teamsUrl) { window.open(teamsUrl, "_blank"); } else { showICSMeetingDetail(m, i, row); }
    };
    list.appendChild(row);
  });
}

let _icsMeetings = [];
window.importICS = function (input) {
  const file = input.files[0];
  if (!file) return;
  document.getElementById("ics-file-name").textContent = file.name;
  const reader = new FileReader();
  reader.onload = (e) => {
    _icsMeetings = parseICS(e.target.result);
    renderICSMeetings(_icsMeetings);
    // Also sync to main calendar
    _icsMeetings.forEach((m) => {
      const k = `${m.start.getFullYear()}-${String(m.start.getMonth()+1).padStart(2,"0")}-${String(m.start.getDate()).padStart(2,"0")}`;
      if (!events[k]) events[k] = [];
      if (!events[k].some((e) => e.title === m.summary)) {
        events[k].push({ title: m.summary, time: m.start.toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit"}), dur: m.dur, att: m.attendees.slice(0,4).join(", "), type: "meet", h: m.start.getHours(), m: m.start.getMinutes() });
      }
    });
    renderCalendar(events);
    renderSchedule(todayKey(), events);
    renderOvMeetings(events);
    document.getElementById("cal-badge").textContent = (events[todayKey()] || []).length;
    showToast(`Imported ${_icsMeetings.length} meeting(s) from ${file.name}`);
  };
  reader.readAsText(file);
};

window.clearICSMeetings = function () {
  _icsMeetings = [];
  document.getElementById("ics-file-name").textContent = "No file chosen";
  document.getElementById("ics-file-input").value = "";
  const countEl = document.getElementById("meetings-count");
  if (countEl) countEl.style.display = "none";
  document.getElementById("teams-list").innerHTML = '<div style="font-size:12px;color:var(--hint);padding:20px 0;text-align:center"><i class="ti ti-calendar-import" style="display:block;font-size:24px;margin-bottom:8px;color:var(--border)"></i>Import your Outlook calendar above to see scheduled meetings.</div>';
  document.getElementById("meeting-detail").innerHTML = "Select a meeting to preview details and generate MOM.";
  const out = document.getElementById("mom-tm-output"); if (out) out.style.display = "none";
};

function showICSMeetingDetail(m, i, row) {
  document.querySelectorAll("#teams-list .meet-row").forEach((r) => r.classList.remove("sel"));
  row.classList.add("sel");
  const det      = document.getElementById("meeting-detail");
  const dateStr  = m.start.toLocaleDateString("en-IN", { weekday:"long", day:"2-digit", month:"long", year:"numeric" });
  const timeStr  = m.start.toLocaleTimeString("en-IN", { hour:"2-digit", minute:"2-digit" });
  det.innerHTML  = `
    <div style="font-size:14px;font-weight:500;margin-bottom:5px">${escHtml(m.summary)}</div>
    <div style="font-size:12px;color:var(--muted);margin-bottom:4px"><i class="ti ti-calendar" style="font-size:11px"></i> ${dateStr} at ${timeStr} · ${m.dur}</div>
    ${m.location ? `<div style="font-size:12px;color:var(--muted);margin-bottom:4px"><i class="ti ti-map-pin" style="font-size:11px"></i> ${escHtml(m.location)}</div>` : ""}
    ${m.attendees.length ? `<div style="font-size:12px;color:var(--muted);margin-bottom:10px"><i class="ti ti-users" style="font-size:11px"></i> ${escHtml(m.attendees.join(", "))}</div>` : ""}
    <button class="primary" onclick="generateICSMeetingMOM(${i})"><i class="ti ti-bolt"></i> Generate MOM</button>`;
}

window.generateICSMeetingMOM = async function (i) {
  const m   = _icsMeetings[i];
  const out = document.getElementById("mom-tm-output");
  const txt = document.getElementById("mom-tm-text");
  if (out) out.style.display = "block";
  if (txt) txt.innerHTML = `<span class="ldot" style="background:var(--blue)"></span>Generating MOM with Groq...`;
  const data = await api.teamsMOM({
    subject:   m.summary,
    date:      m.start.toLocaleDateString("en-IN", { day:"2-digit", month:"long", year:"numeric" }),
    attendees: m.attendees.join(", ") || "Not specified",
    duration:  m.dur,
    context:   m.desc || m.location || "",
  });
  if (txt) typeIn(txt, data.text || data.error || "Error generating MOM");
};

window.downloadMOM = function () {
  const txt = document.getElementById("mom-tm-text")?.innerText || "";
  if (!txt) return;
  const a   = document.createElement("a");
  a.href    = URL.createObjectURL(new Blob([txt], { type: "text/plain" }));
  a.download = "MOM.txt";
  a.click();
};

// ── Exchange Web Services (automatic on-premise sync) ──────────
let _ewsAutoSyncTimer = null;

function ewsSaveCreds(url, username, password) {
  localStorage.setItem("ewsUrl",      url);
  localStorage.setItem("ewsUsername", username);
  localStorage.setItem("ewsPassword", password);
}
function ewsGetCreds() {
  return {
    ewsUrl:   localStorage.getItem("ewsUrl")      || "",
    username: localStorage.getItem("ewsUsername") || "",
    password: localStorage.getItem("ewsPassword") || "",
  };
}
function extractFirstName(username) {
  const local = username.includes("@") ? username.split("@")[0] : username;
  const first = local.split(/[.\-_]/)[0];
  return first.charAt(0).toUpperCase() + first.slice(1).toLowerCase();
}

window.updateWelcomeMessage = function () {
  const firstName = localStorage.getItem("firstName") || "";
  const el = document.getElementById("tb-welcome");
  if (!el) return;
  if (firstName) { el.textContent = `Welcome, ${firstName}`; el.style.display = "block"; }
  else { el.style.display = "none"; }
};

function showOvMain() {
  const el = document.getElementById("ov-main-content");
  if (el) el.style.display = "block";
  const prompt = document.getElementById("ov-setup-prompt");
  if (prompt) prompt.style.display = "none";
  const hint = document.getElementById("ov-setup-hint");
  if (hint) hint.style.display = "none";
}

function ewsRestoreUI() {
  const { ewsUrl, username } = ewsGetCreds();
  const setupEl     = document.getElementById("ews-setup-form");
  const connectedEl = document.getElementById("ews-connected-state");
  const userLbl     = document.getElementById("ews-user-label");

  if (ewsUrl && username) {
    if (setupEl)     setupEl.style.display     = "none";
    if (connectedEl) connectedEl.style.display = "block";
    if (userLbl) userLbl.textContent = username + "  ·  " + ewsUrl.replace("https://","").replace("/EWS/Exchange.asmx","");
    const olHint = document.getElementById("ol-ews-hint");
    const olOk   = document.getElementById("ol-ews-ok");
    if (olHint) olHint.style.display = "none";
    if (olOk)   olOk.style.display   = "inline";
  } else {
    if (setupEl)     setupEl.style.display     = "block";
    if (connectedEl) connectedEl.style.display = "none";
  }
}

function renderEWSMeetings(meetings) {
  const list    = document.getElementById("teams-list");
  const countEl = document.getElementById("meetings-count");
  if (!list) return;
  if (!meetings || !meetings.length) {
    list.innerHTML = '<div style="font-size:12px;color:var(--hint);padding:20px 0;text-align:center"><i class="ti ti-calendar-off" style="display:block;font-size:24px;margin-bottom:8px;color:var(--border)"></i>No meetings found.</div>';
    if (countEl) countEl.style.display = "none";
    return;
  }
  if (countEl) { countEl.style.display = "inline"; countEl.textContent = meetings.length; }

  // Show today + past 5 days only — no upcoming future meetings
  const now        = new Date();
  const todayStart = new Date(now); todayStart.setHours(0,0,0,0);
  const todayEnd   = new Date(now); todayEnd.setHours(23,59,59,999);
  const fiveDaysAgo = new Date(now); fiveDaysAgo.setDate(fiveDaysAgo.getDate() - 5); fiveDaysAgo.setHours(0,0,0,0);

  const todayMtgs = meetings
    .filter(m => { const d = new Date(m.start); return d >= todayStart && d <= todayEnd; })
    .sort((a, b) => new Date(a.start) - new Date(b.start));
  const past = meetings
    .filter(m => { const d = new Date(m.start); return d < todayStart && d >= fiveDaysAgo; })
    .sort((a, b) => new Date(b.start) - new Date(a.start));
  const sorted = [...todayMtgs, ...past];
  _sortedMeetings = sorted;

  list.innerHTML = "";
  sorted.forEach((m, i) => {
    const start    = new Date(m.start);
    const isPast   = start < now;
    const isToday  = start.toDateString() === now.toDateString();
    const isDone   = isPast || isToday;
    const dateStr  = isNaN(start) ? m.start : start.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
    const timeStr  = isNaN(start) ? "" : start.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" });
    const atts     = (m.attendees || []).slice(0, 3).join(", ");
    const row      = document.createElement("div");
    row.className  = "meet-row" + (isToday ? " sel" : "");
    row.style.alignItems = "center";
    row.innerHTML  = `
      <div style="flex:1;min-width:0">
        <div class="mr-t">${escHtml(m.subject)}</div>
        <div class="mr-m">${dateStr}${timeStr ? " · " + timeStr : ""} · ${m.dur || ""}${atts ? "  ·  " + escHtml(atts) : ""}</div>
      </div>
      <span class="badge ${isToday ? "b-amber" : "b-green"}" style="flex-shrink:0">${isToday ? "Today" : "Past"}</span>
      ${m.isOnline ? '<span class="badge b-blue" style="margin-left:4px;flex-shrink:0">Teams</span>' : ""}`;
    row.onclick = () => {
      if (m.joinUrl) { window.open(m.joinUrl, "_blank"); } else { showEWSMeetingDetail(m, i, row); }
    };
    list.appendChild(row);
  });

  // sync meetings to calendar
  meetings.forEach((m) => {
    const start = new Date(m.start);
    if (isNaN(start)) return;
    const k = `${start.getFullYear()}-${String(start.getMonth()+1).padStart(2,"0")}-${String(start.getDate()).padStart(2,"0")}`;
    if (!events[k]) events[k] = [];
    if (!events[k].some((e) => e.title === m.subject)) {
      events[k].push({ title: m.subject, time: start.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" }), dur: m.dur || "N/A", att: (m.attendees || []).slice(0, 4).join(", "), type: "meet", h: start.getHours(), min: start.getMinutes() });
    }
  });
  renderCalendar(events);
  renderSchedule(todayKey(), events);
  renderOvMeetings(events);
  document.getElementById("cal-badge").textContent = (events[todayKey()] || []).length;
  if (window.populateMOMDropdown) window.populateMOMDropdown();
}

let _ewsMeetings = [];
let _sortedMeetings = [];
const _meetingTranscripts = new Map(); // index → transcript text saved during meeting

function showEWSMeetingDetail(m, i, row) {
  document.querySelectorAll("#teams-list .meet-row").forEach((r) => r.classList.remove("sel"));
  row.classList.add("sel");
  const det     = document.getElementById("meeting-detail");
  if (!det) return;

  const start    = new Date(m.start);
  const dateStr  = isNaN(start) ? m.start : start.toLocaleDateString("en-IN", { weekday: "long", day: "2-digit", month: "long", year: "numeric" });
  const timeStr  = isNaN(start) ? "" : start.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" });
  const joinUrl  = m.joinUrl || "";
  const saved    = _meetingTranscripts.get(i) || "";

  det.innerHTML = `
    <div style="font-size:14px;font-weight:500;color:var(--text);margin-bottom:6px">${escHtml(m.subject)}</div>
    <div style="font-size:12px;color:var(--muted);margin-bottom:3px"><i class="ti ti-calendar" style="font-size:11px"></i> ${dateStr}${timeStr ? " at " + timeStr : ""} · ${m.dur || "N/A"}</div>
    ${m.location ? `<div style="font-size:12px;color:var(--muted);margin-bottom:3px"><i class="ti ti-map-pin" style="font-size:11px"></i> ${escHtml(m.location)}</div>` : ""}
    ${m.attendees && m.attendees.length ? `<div style="font-size:12px;color:var(--muted);margin-bottom:10px"><i class="ti ti-users" style="font-size:11px"></i> ${escHtml(m.attendees.join(", "))}</div>` : ""}

    ${joinUrl
      ? `<a href="${escHtml(joinUrl)}" target="_blank" style="text-decoration:none"><button class="primary" style="width:100%;margin-bottom:10px"><i class="ti ti-video"></i> Join Teams Meeting</button></a>`
      : m.isOnline
        ? `<div style="font-size:12px;color:var(--muted);margin-bottom:10px;padding:7px 10px;background:var(--surface2);border-radius:var(--r-sm);border:1px solid var(--border)"><i class="ti ti-video" style="font-size:11px"></i> Online meeting — join link will appear after next sync if Exchange returns it.</div>`
        : ""}

    <div style="margin-bottom:6px">
      <label style="font-size:11px;font-weight:500;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);margin-bottom:5px;display:block"><i class="ti ti-microphone" style="font-size:11px"></i> Live notes / transcript</label>
      <textarea id="mt-area-${i}" style="min-height:100px;font-size:12px;margin-bottom:8px;resize:vertical" placeholder="Type or paste notes while the meeting is in progress…" oninput="_saveMeetingTranscript(${i}, this.value)">${escHtml(saved)}</textarea>
    </div>
    <button class="primary" onclick="generateMOMWithTranscript(${i})"><i class="ti ti-notes"></i> Generate MOM from notes</button>`;
}

window._saveMeetingTranscript = function (i, text) {
  _meetingTranscripts.set(i, text);
};

window.generateMOMWithTranscript = async function (i) {
  const m = _sortedMeetings[i] || _ewsMeetings[i];
  if (!m) return;
  const transcript = _meetingTranscripts.get(i) || document.getElementById(`mt-area-${i}`)?.value?.trim() || "";
  const out = document.getElementById("mom-tm-output");
  const txt = document.getElementById("mom-tm-text");
  if (out) out.style.display = "block";
  if (txt) txt.innerHTML = `<span class="ldot" style="background:var(--blue)"></span>Generating MOM${transcript ? " from transcript" : ""}...`;
  const start = new Date(m.start);
  const data  = await api.teamsMOM({
    subject:    m.subject,
    date:       isNaN(start) ? m.start : start.toLocaleDateString("en-IN", { day: "2-digit", month: "long", year: "numeric" }),
    attendees:  (m.attendees || []).join(", ") || "Not specified",
    duration:   m.dur || "N/A",
    context:    m.location || "",
    transcript: transcript || undefined,
  });
  if (txt) typeIn(txt, data.text || data.error || "Error generating MOM");
};

// Keep legacy name for backward compat with ICS path
window.generateEWSMeetingMOM = window.generateMOMWithTranscript;

window.openMOMModal = function (i) {
  const m = _sortedMeetings[i];
  if (!m) return;
  const start   = new Date(m.start);
  const dateStr = isNaN(start) ? m.start : start.toLocaleDateString("en-IN", { weekday: "short", day: "2-digit", month: "short", year: "numeric" });
  const timeStr = isNaN(start) ? "" : start.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" });
  const atts    = (m.attendees || []).join(", ") || "—";

  const overlay = document.createElement("div");
  overlay.style.cssText = "position:fixed;inset:0;z-index:10000;background:rgba(0,0,0,.55);display:flex;align-items:center;justify-content:center;padding:20px";

  const modal = document.createElement("div");
  modal.style.cssText = "background:var(--surface);border:1px solid var(--border);border-radius:var(--r);max-width:700px;width:100%;max-height:90vh;overflow-y:auto;padding:22px;box-shadow:0 8px 40px rgba(0,0,0,.22)";

  modal.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:start;margin-bottom:14px">
      <div>
        <div style="font-size:15px;font-weight:600;color:var(--text);margin-bottom:4px">${escHtml(m.subject)}</div>
        <div style="font-size:12px;color:var(--muted)">${dateStr}${timeStr ? " · " + timeStr : ""} · ${m.dur || "N/A"}</div>
        <div style="font-size:12px;color:var(--muted)"><i class="ti ti-users" style="font-size:11px"></i> ${escHtml(atts)}</div>
      </div>
      <button class="sm" id="mom-modal-close"><i class="ti ti-x"></i></button>
    </div>
    <label style="font-size:11px;font-weight:500;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);margin-bottom:5px;display:block">Meeting transcript / notes <span style="font-weight:400;color:var(--hint)">(paste from call recording or type notes)</span></label>
    <textarea id="mom-modal-transcript" style="min-height:110px;font-size:12px;margin-bottom:10px" placeholder="Paste meeting transcript, key discussion points, or notes here. If left blank, MOM will be generated from meeting metadata."></textarea>
    <button class="primary" id="mom-modal-gen-btn" onclick="generateMOMFromModal(${i}, this)"><i class="ti ti-bolt"></i> Generate MOM</button>
    <div id="mom-modal-output" style="display:none;margin-top:14px">
      <div class="term" id="mom-modal-text" style="min-height:120px"></div>
      <div class="btn-row" style="margin-top:8px">
        <button onclick="copyText('mom-modal-text')"><i class="ti ti-copy"></i> Copy</button>
        <button onclick="transferMOMToPanel()"><i class="ti ti-arrow-right"></i> Open in MOM panel</button>
      </div>
    </div>`;

  overlay.appendChild(modal);
  document.body.appendChild(overlay);
  modal.querySelector("#mom-modal-close").onclick = () => overlay.remove();
  overlay.onclick = (e) => { if (e.target === overlay) overlay.remove(); };
};

window.generateMOMFromModal = async function (i, btn) {
  const m          = _sortedMeetings[i];
  const transcript = document.getElementById("mom-modal-transcript")?.value?.trim() || "";
  const outDiv     = document.getElementById("mom-modal-output");
  const txt        = document.getElementById("mom-modal-text");
  if (!m || !outDiv || !txt) return;

  outDiv.style.display = "block";
  txt.innerHTML = `<span class="ldot" style="background:var(--blue)"></span>Generating MOM...`;
  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader" style="font-size:12px;animation:spin .7s linear infinite"></i> Generating...'; }

  const start = new Date(m.start);
  const data  = await api.teamsMOM({
    subject:    m.subject,
    date:       isNaN(start) ? m.start : start.toLocaleDateString("en-IN", { day: "2-digit", month: "long", year: "numeric" }),
    attendees:  (m.attendees || []).join(", ") || "Not specified",
    duration:   m.dur || "N/A",
    context:    m.location || "",
    transcript: transcript || undefined,
  });

  if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-bolt"></i> Generate MOM'; }
  typeIn(txt, data.text || data.error || "Error generating MOM");
};

window.transferMOMToPanel = function () {
  const txt = document.getElementById("mom-modal-text")?.innerText || "";
  if (!txt) return;
  const panelTxt = document.getElementById("mom-tm-text");
  const panelOut = document.getElementById("mom-tm-output");
  if (panelTxt) panelTxt.textContent = txt;
  if (panelOut) panelOut.style.display = "block";
  showToast("MOM transferred to panel.");
  document.querySelector(".email-modal-overlay, [style*='z-index:10000']")?.remove();
};

window.ewsCheckInputs = function () {
  const username = document.getElementById("ews-username")?.value?.trim();
  const password = document.getElementById("ews-password")?.value?.trim();
  const btn = document.getElementById("ews-connect-btn");
  if (btn) btn.disabled = !(username && password);
};

window.ewsConnect = async function (btn) {
  const username = document.getElementById("ews-username")?.value?.trim();
  const password = document.getElementById("ews-password")?.value?.trim();
  const urlEl    = document.getElementById("ews-url");
  const errEl    = document.getElementById("ews-error");

  if (!username || !password) {
    if (errEl) { errEl.textContent = "Please enter your username and password."; errEl.style.display = "block"; }
    return;
  }
  if (errEl) errEl.style.display = "none";

  // Use the field value; if blank, auto-derive from email domain
  let url = urlEl?.value?.trim();
  if (!url) {
    const domain = username.includes("@") ? username.split("@")[1] : "";
    if (!domain) { if (errEl) { errEl.textContent = "Enter the Exchange server URL or use Auto-detect."; errEl.style.display = "block"; } return; }
    url = `https://owa.${domain}/EWS/Exchange.asmx`;
    if (urlEl) urlEl.value = url;
  }

  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader" style="font-size:12px;animation:spin .7s linear infinite"></i> Connecting…'; }

  const data = await api.ewsMeetings({ ewsUrl: url, username, password });

  if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-plug"></i> Connect &amp; Sync'; }

  if (data.error) {
    if (errEl) { errEl.textContent = data.error; errEl.style.display = "block"; }
    return;
  }

  ewsSaveCreds(url, username, password);

  // Save first name for welcome message
  const firstName = extractFirstName(username);
  localStorage.setItem("firstName", firstName);
  window.updateWelcomeMessage();

  // Remove login-mode so sidebar + topbar reappear
  document.getElementById("app")?.classList.remove("login-mode");
  stopTmBackground();
  ewsRestoreUI();

  _ewsMeetings = data.meetings || [];
  renderEWSMeetings(_ewsMeetings);
  updateOverviewStats(_ewsMeetings, null);
  renderOvUpcoming(_ewsMeetings);

  showToast(`Connected! Loaded ${_ewsMeetings.length} meeting(s).`);

  // Redirect to Dashboard — show full content only if project already set
  navTo("ov");
  if (localStorage.getItem("projectName")) {
    showOvMain();
  }
};

window.ewsSync = async function (btn) {
  const creds  = ewsGetCreds();
  if (!creds.ewsUrl) { showToast("Not connected. Fill in the server details first."); return; }

  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader" style="font-size:12px;animation:spin .7s linear infinite"></i>'; }

  const data = await api.ewsMeetings(creds);

  if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-refresh"></i> Sync now'; }

  if (data.error) { showToast("EWS sync error: " + data.error); return; }

  _ewsMeetings = data.meetings || [];
  renderEWSMeetings(_ewsMeetings);
  updateOverviewStats(_ewsMeetings, null);
  renderOvUpcoming(_ewsMeetings);

  const syncLbl = document.getElementById("ews-last-sync-lbl");
  if (syncLbl) syncLbl.textContent = "Last synced: " + new Date().toLocaleTimeString("en-IN");

  showToast(`Synced ${_ewsMeetings.length} meeting(s) from Exchange.`);
};

window.ewsDisconnect = function () {
  localStorage.removeItem("ewsUrl");
  localStorage.removeItem("ewsUsername");
  localStorage.removeItem("ewsPassword");
  if (_ewsAutoSyncTimer) { clearInterval(_ewsAutoSyncTimer); _ewsAutoSyncTimer = null; }
  _ewsMeetings = [];
  const countEl = document.getElementById("meetings-count");
  if (countEl) countEl.style.display = "none";
  document.getElementById("app")?.classList.add("login-mode");
  // Clear welcome + reset setup state for next login
  localStorage.removeItem("firstName");
  localStorage.removeItem("projectName");
  window.updateWelcomeMessage();
  const ovMain = document.getElementById("ov-main-content");
  if (ovMain) ovMain.style.display = "none";
  const ovPrompt = document.getElementById("ov-setup-prompt");
  if (ovPrompt) ovPrompt.style.display = "flex";
  const sbProj = document.getElementById("sb-project-name");
  if (sbProj) { sbProj.textContent = ""; sbProj.style.display = "none"; }
  const projInp = document.getElementById("ov-project");
  if (projInp) projInp.value = "";
  ewsRestoreUI();
  navTo("tm");
  startTmBackground();
  showToast("Signed out from Exchange.");
};

window.ewsToggleAutoSync = function (btn) {
  const active = btn.dataset.active === "1";
  if (!active) {
    _ewsAutoSyncTimer = setInterval(() => window.ewsSync(null), 5 * 60 * 1000);
    btn.dataset.active = "1";
    btn.innerHTML = '<i class="ti ti-refresh"></i> Auto-sync on';
    showToast("Auto-sync enabled — refreshes every 5 minutes.");
  } else {
    if (_ewsAutoSyncTimer) { clearInterval(_ewsAutoSyncTimer); _ewsAutoSyncTimer = null; }
    btn.dataset.active = "0";
    btn.innerHTML = '<i class="ti ti-refresh"></i> Auto-sync off';
    showToast("Auto-sync disabled.");
  }
};

window.ewsDiscoverScrollOff = function ewsDiscoverScrollOff() {
  const content = document.getElementById("content");
  const ptm     = document.getElementById("p-tm");
  if (content) { content.style.overflowY = "hidden"; content.style.height = "100vh"; }
  if (ptm)     { ptm.style.overflow = "hidden"; ptm.style.height = "100vh"; ptm.style.minHeight = ""; }
}

function ewsDiscoverScrollOn() {
  const content = document.getElementById("content");
  const ptm     = document.getElementById("p-tm");
  if (content) { content.style.overflowY = "auto"; content.style.height = "100vh"; }
  if (ptm)     { ptm.style.overflow = "visible"; ptm.style.height = "auto"; ptm.style.minHeight = "100vh"; }
}
document.addEventListener("click", (e) => {
  const res = document.getElementById("ews-discover-results");
  if (res && res.style.display !== "none" &&
      !res.contains(e.target) && e.target.id !== "ews-discover-btn" &&
      !e.target.closest("#ews-discover-btn")) {
    res.style.display = "none";
    ewsDiscoverScrollOff();
  }
});

window.ewsDiscover = async function (btn) {
  const username = document.getElementById("ews-username")?.value?.trim();
  const email    = username;
  if (!email) { showToast("Enter your email first, then click Auto-detect."); if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-search"></i> Auto-detect'; } return; }
  const resDiv   = document.getElementById("ews-discover-results");
  if (!resDiv) return;
  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader" style="font-size:12px;animation:spin .7s linear infinite"></i> Detecting...'; }

  const data = await api.ewsDiscover(email);

  if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-search"></i> Auto-detect'; }

  if (data.error) { showToast(data.error); return; }

  const candidates = data.candidates || [];
  resDiv.style.display = "block";
  ewsDiscoverScrollOn();
  resDiv.innerHTML = `<div style="font-size:12px;color:#93c5fd;margin-bottom:6px">Common EWS URLs for <strong style="color:#fff">${escHtml(data.domain)}</strong> — click one to use it:</div>` +
    candidates.map((c) => `<div style="padding:5px 0;border-bottom:1px solid rgba(99,179,237,0.2);font-size:12px;font-family:var(--mono);display:flex;align-items:center;justify-content:space-between;gap:8px"><span style="color:#e2e8f0;word-break:break-all">${escHtml(c)}</span><button class="sm" style="flex-shrink:0" onclick="document.getElementById('ews-url').value='${escHtml(c)}';document.getElementById('ews-discover-results').style.display='none';ewsDiscoverScrollOff()">Use</button></div>`).join("");
};

// ── Teams MOM — uses live EWS meetings (_sortedMeetings) ───────
window.generateTeamsMOM = async function (i) {
  const m   = _sortedMeetings[i] || _ewsMeetings[i];
  if (!m) return;
  const out = document.getElementById("mom-tm-output");
  const txt = document.getElementById("mom-tm-text");
  if (out) out.style.display = "block";
  if (txt) txt.innerHTML = `<span class="ldot" style="background:var(--blue)"></span>Generating MOM…`;
  const start   = new Date(m.start);
  const dateStr = isNaN(start) ? m.start : start.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
  const atts    = (m.attendees || []).join(", ");
  const data = await api.generateMom({
    title:     m.subject,
    date:      dateStr,
    attendees: atts,
    objective: m.location || m.joinUrl || "See calendar invite",
    provider:  localStorage.getItem("ai_provider") || "groq",
  });
  if (txt) typeIn(txt, data.text || data.error || "Error generating MOM");
};

// ── Teams — sync real meetings from Microsoft 365 ──────────────
window.syncTeamsMeetings = async function (btn) {
  const list = document.getElementById("teams-list");
  if (!getMsToken()) {
    showToast("Paste your Graph Explorer token first and click Connect.");
    return;
  }
  btn.disabled = true;
  btn.innerHTML = '<i class="ti ti-loader" style="font-size:12px;animation:spin .7s linear infinite"></i> Syncing...';
  const data = await api.teamsMeetings();
  btn.disabled = false;
  btn.innerHTML = '<i class="ti ti-cloud-download"></i> Sync from M365';
  if (data.hint === "permission") {
    list.innerHTML = `<div style="font-size:12px;color:var(--amber);padding:12px;background:var(--surface2);border-radius:var(--r-sm);border:1px solid var(--border)">${data.error}</div>`;
    return;
  }
  if (data.onPremise) {
    list.innerHTML = `<div style="font-size:12px;color:var(--muted);padding:12px;background:var(--surface2);border-radius:var(--r-sm);border:1px solid var(--border)"><strong style="color:var(--text)">On-premise Exchange detected.</strong><br>${data.error}<br><br>You can still use the <strong>MOM Generator</strong> manually to generate meeting minutes.</div>`;
    return;
  }
  if (data.error) { showToast("Teams sync error: " + data.error); return; }
  if (!data.meetings?.length) { showToast("No Teams meetings found in the next 14 days."); return; }
  list.innerHTML = "";
  data.meetings.forEach((m, i) => {
    const start = new Date(m.start?.dateTime || "");
    const end   = new Date(m.end?.dateTime   || "");
    const dur   = isNaN(start) ? "" : `${Math.round((end - start) / 60000)} min`;
    const atts  = (m.attendees || []).map((a) => a.emailAddress?.name || a.emailAddress?.address).slice(0, 4).join(", ");
    const row   = document.createElement("div");
    row.className = "meet-row";
    row.onclick   = () => {
      if (m.onlineMeetingUrl) { window.open(m.onlineMeetingUrl, "_blank"); } else { selectMeeting(row, -1, m); }
    };
    row.innerHTML = `<div><div class="mr-t">${escHtml(m.subject || "Untitled")}</div><div class="mr-m">${isNaN(start) ? "" : start.toLocaleDateString("en-IN", { day: "2-digit", month: "short" })} · ${dur} · ${escHtml(atts)}</div></div><span class="badge b-blue">M365</span>`;
    list.appendChild(row);
  });
  showToast(`Loaded ${data.meetings.length} Teams meeting(s).`);
};

// Override selectMeeting to support real M365 meetings
const _origSelectMeeting = window.selectMeeting;
window.selectMeeting = function (el, i, msData) {
  document.querySelectorAll(".meet-row").forEach((r) => r.classList.remove("sel"));
  el.classList.add("sel");
  const det = document.getElementById("meeting-detail");
  if (!det) return;
  if (msData) {
    const start = new Date(msData.start?.dateTime || "");
    const end   = new Date(msData.end?.dateTime   || "");
    const dur   = isNaN(start) ? "N/A" : `${Math.round((end - start) / 60000)} min`;
    const atts  = (msData.attendees || []).map((a) => a.emailAddress?.name || a.emailAddress?.address).join(", ");
    det.innerHTML = `<div style="font-size:14px;font-weight:500;margin-bottom:5px">${escHtml(msData.subject || "Meeting")}</div><div style="font-size:12px;color:var(--muted);margin-bottom:7px">${isNaN(start) ? "" : start.toLocaleString("en-IN")} · ${dur} · ${escHtml(atts)}</div><button class="primary" onclick="generateTeamsMOMFromGraph('${encodeURIComponent(JSON.stringify({ subject: msData.subject, date: start.toLocaleDateString('en-IN'), attendees: atts, duration: dur }))}')"><i class="ti ti-bolt"></i> Generate MOM</button>`;
  } else {
    _origSelectMeeting(el, i);
  }
};
window.generateTeamsMOMFromGraph = async function (encoded) {
  const m   = JSON.parse(decodeURIComponent(encoded));
  const out = document.getElementById("mom-tm-output");
  const txt = document.getElementById("mom-tm-text");
  if (out) out.style.display = "block";
  if (txt) txt.innerHTML = `<span class="ldot" style="background:var(--blue)"></span>Generating MOM...`;
  const data = await api.teamsMOM(m);
  if (txt) typeIn(txt, data.text || data.error || "Error generating MOM");
};

// ── Outlook inbox via EWS ──────────────────────────────────────
let _currentEmail = null;

window.syncEWSEmails = async function (btn) {
  const list  = document.getElementById("ol-list");
  const creds = ewsGetCreds();
  if (!creds.ewsUrl) {
    list.innerHTML = '<div style="font-size:12px;color:var(--muted);padding:14px;background:var(--surface2);border-radius:var(--r-sm)">Connect to Exchange in the <strong>Teams meetings</strong> panel first.</div>';
    return;
  }
  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader" style="font-size:12px;animation:spin .7s linear infinite"></i> Loading...'; }

  const data = await api.ewsEmails(creds);

  if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-refresh"></i> Load inbox'; }

  const hint = document.getElementById("ol-ews-hint");
  const ok   = document.getElementById("ol-ews-ok");

  if (data.error) {
    if (hint) hint.style.display = "inline";
    if (ok)   ok.style.display   = "none";
    list.innerHTML = `<div style="font-size:12px;color:var(--red);padding:12px;background:rgba(220,38,38,.06);border-radius:var(--r-sm);border:1px solid rgba(220,38,38,.15)">${data.error}</div>`;
    return;
  }

  if (hint) hint.style.display = "none";
  if (ok)   ok.style.display   = "inline";

  const emails = data.emails || [];
  const badge  = document.getElementById("ol-badge");
  const unread = emails.filter((e) => !e.isRead).length;
  if (badge) { badge.textContent = unread; badge.style.display = unread ? "inline" : "none"; }
  updateOverviewStats(_ewsMeetings, unread);

  if (!emails.length) {
    list.innerHTML = '<div style="font-size:12px;color:var(--hint);padding:16px 0;text-align:center">No emails in inbox.</div>';
    return;
  }

  list.innerHTML = "";
  emails.forEach((email) => {
    const row  = document.createElement("div");
    row.className = "meet-row";
    row.style.cursor = "pointer";
    const from = email.from?.name || email.from?.address || "Unknown";
    const date = email.received
      ? new Date(email.received).toLocaleDateString("en-IN", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" })
      : "";
    row.innerHTML = `
      <div style="flex:1;min-width:0">
        <div class="mr-t" style="${!email.isRead ? "font-weight:600" : ""}">${escHtml(email.subject)}</div>
        <div class="mr-m">${escHtml(from)} · ${date}</div>
      </div>
      ${!email.isRead ? '<span class="badge b-blue">Unread</span>' : ""}`;
    row.onclick    = () => previewEWSEmail(email, row, creds);
    row.ondblclick = () => openEmailExpanded(email, creds);
    list.appendChild(row);
  });
  showToast(`Loaded ${emails.length} email(s) from Exchange.`);
};

async function previewEWSEmail(email, row, creds) {
  document.querySelectorAll("#ol-list .meet-row").forEach((r) => r.classList.remove("sel"));
  row.classList.add("sel");
  _currentEmail = email;
  const det = document.getElementById("ol-detail");
  if (!det) return;

  const from        = email.from?.name || email.from?.address || "Unknown";
  const fromAddress = email.from?.address || "";
  const dateStr     = email.received ? new Date(email.received).toLocaleString("en-IN") : "";

  det.innerHTML = `
    <div style="border-bottom:1px solid var(--border);padding-bottom:10px;margin-bottom:10px">
      <div style="font-size:15px;font-weight:600;color:var(--text);margin-bottom:5px">${escHtml(email.subject || "(no subject)")}</div>
      <div style="font-size:12px;color:var(--muted)"><strong>From:</strong> ${escHtml(from)}${fromAddress ? " &lt;" + escHtml(fromAddress) + "&gt;" : ""}</div>
      <div style="font-size:12px;color:var(--muted)"><strong>Date:</strong> ${dateStr}</div>
    </div>
    <div id="ol-body-content" style="min-height:60px;margin-bottom:12px;border-radius:var(--r-sm);overflow:hidden">
      <div style="font-size:12px;color:var(--muted);padding:10px 0"><span class="ldot" style="background:var(--blue)"></span> Loading full email...</div>
    </div>
    <div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:10px">
      <button class="sm primary" onclick="startEmailCompose('reply')"><i class="ti ti-arrow-back-up"></i> Reply</button>
      <button class="sm" onclick="startEmailCompose('forward')"><i class="ti ti-arrow-forward-up"></i> Forward</button>
      <button class="sm" onclick="draftEmailReply()"><i class="ti ti-robot"></i> AI Draft</button>
    </div>
    <div id="ol-compose-area" style="display:none;border-top:1px solid var(--border);padding-top:10px">
      <label style="font-size:11px;font-weight:500;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);margin-bottom:4px;display:block">To</label>
      <input id="ol-compose-to" style="margin-bottom:6px;font-size:12px" placeholder="recipient@email.com">
      <label style="font-size:11px;font-weight:500;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);margin-bottom:4px;display:block">Subject</label>
      <input id="ol-compose-subject" style="margin-bottom:6px;font-size:12px">
      <label style="font-size:11px;font-weight:500;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);margin-bottom:4px;display:block">Message</label>
      <textarea id="ol-compose-body" style="min-height:90px;font-size:12px;margin-bottom:8px"></textarea>
      <div style="display:flex;gap:6px">
        <button class="primary sm" onclick="sendComposedEmail(this)"><i class="ti ti-send"></i> Send</button>
        <button class="sm" onclick="document.getElementById('ol-compose-area').style.display='none'">Cancel</button>
      </div>
    </div>`;

  const draftSec = document.getElementById("ol-draft-section");
  if (draftSec) draftSec.style.display = "none";

  if (email.id && creds) {
    const bodyData = await api.ewsEmailBody({ ...creds, itemId: email.id, changeKey: email.changeKey });
    const bodyEl   = document.getElementById("ol-body-content");
    const content  = bodyData.body || "(no content)";
    const isHtml   = bodyData.bodyType === "html" || /<[a-z][\s\S]*>/i.test(content);
    if (bodyEl) {
      if (isHtml) {
        const iframe = document.createElement("iframe");
        iframe.sandbox = "allow-same-origin allow-popups";
        iframe.style.cssText = "width:100%;border:none;min-height:300px;border-radius:var(--r-sm);background:#fff";
        iframe.srcdoc = content;
        iframe.onload = () => {
          const h = iframe.contentDocument?.body?.scrollHeight;
          if (h) iframe.style.height = h + 20 + "px";
        };
        bodyEl.innerHTML = "";
        bodyEl.appendChild(iframe);
      } else {
        bodyEl.style.cssText = "font-size:13px;color:var(--text);white-space:pre-wrap;line-height:1.7;padding:8px 0;margin-bottom:12px";
        bodyEl.textContent = content;
      }
    }
    _currentEmail.bodyFull    = content;
    _currentEmail.bodyPreview = content.replace(/<[^>]+>/g, "").slice(0, 500);
  }
}

window.startEmailCompose = function (type) {
  if (!_currentEmail) return;
  const area    = document.getElementById("ol-compose-area");
  const toEl    = document.getElementById("ol-compose-to");
  const subEl   = document.getElementById("ol-compose-subject");
  const bodyEl  = document.getElementById("ol-compose-body");
  if (!area || !toEl || !subEl || !bodyEl) return;

  const fromAddr = _currentEmail.from?.address || "";
  const subject  = _currentEmail.subject || "";
  const body     = _currentEmail.bodyFull || _currentEmail.bodyPreview || "";
  const quote    = "\n\n---------- " + (type === "reply" ? "Original" : "Forwarded") + " message ----------\n" + body.slice(0, 600);

  if (type === "reply") {
    toEl.value  = fromAddr;
    subEl.value = subject.startsWith("Re:") ? subject : "Re: " + subject;
    bodyEl.value = quote;
  } else {
    toEl.value   = "";
    subEl.value  = subject.startsWith("Fwd:") ? subject : "Fwd: " + subject;
    bodyEl.value = "\n\nFrom: " + fromAddr + quote;
  }
  area.style.display = "block";
  toEl.focus();
};

window.showOlCompose = function () {
  const area = document.getElementById("ol-compose-area");
  if (!area) return;
  area.style.display = "block";
  document.getElementById("ol-compose-to")?.focus();
};

window.aiDraftCompose = async function () {
  const subject = document.getElementById("ol-compose-subject")?.value?.trim();
  const to      = document.getElementById("ol-compose-to")?.value?.trim();
  const bodyEl  = document.getElementById("ol-compose-body");
  if (!subject) { showToast("Enter a subject first so the AI can draft the email."); return; }
  if (bodyEl) bodyEl.value = "Drafting…";
  const data = await api.chat(`Write a professional email with subject: "${subject}"${to ? " to " + to : ""}. Include a greeting, clear body, and professional closing. Do not include subject line in the body.`, "You write concise professional workplace emails.");
  if (bodyEl) bodyEl.value = data.text || data.error || "";
};

window.sendComposedEmail = async function (btn) {
  const to      = document.getElementById("ol-compose-to")?.value?.trim();
  const cc      = document.getElementById("ol-compose-cc")?.value?.trim();
  const subject = document.getElementById("ol-compose-subject")?.value?.trim();
  const body    = document.getElementById("ol-compose-body")?.value?.trim();
  const statusEl = document.getElementById("ol-compose-status");

  if (!to)      { showToast("Enter a recipient email."); return; }
  if (!subject) { showToast("Enter a subject."); return; }
  if (!body)    { showToast("Message body is empty."); return; }

  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader" style="font-size:12px;animation:spin .7s linear infinite"></i> Sending…'; }
  if (statusEl) { statusEl.style.display = "block"; statusEl.style.background = "var(--surface2)"; statusEl.style.color = "var(--muted)"; statusEl.textContent = "Sending via Exchange…"; }

  const creds = ewsGetCreds();
  if (!creds.ewsUrl) {
    if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-send"></i> Send Email'; }
    if (statusEl) { statusEl.style.background = "rgba(220,150,0,.1)"; statusEl.style.color = "var(--orange,#d97706)"; statusEl.textContent = "⚠ Not connected to Exchange. Connect in the Teams panel first."; }
    return;
  }

  const data = await api.ewsSendEmail({ ...creds, to, cc, subject, body });

  if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-send"></i> Send Email'; }

  if (data.ok) {
    if (statusEl) { statusEl.style.background = "var(--green-bg)"; statusEl.style.color = "var(--green)"; statusEl.textContent = `✓ Email sent to ${to}`; }
    showToast(`Email sent to ${to}!`);
    setTimeout(() => {
      document.getElementById("ol-compose-to").value      = "";
      document.getElementById("ol-compose-cc").value      = "";
      document.getElementById("ol-compose-subject").value = "";
      document.getElementById("ol-compose-body").value    = "";
      if (statusEl) statusEl.style.display = "none";
    }, 2000);
  } else {
    if (statusEl) { statusEl.style.background = "rgba(220,38,38,.07)"; statusEl.style.color = "var(--red)"; statusEl.textContent = `✗ ${data.error || "Send failed"}`; }
    showToast("Send failed: " + (data.error || "Unknown error"), 5000);
  }
};

async function openEmailExpanded(email, creds) {
  const from    = email.from?.name || email.from?.address || "Unknown";
  const subject = email.subject || "(no subject)";
  const date    = email.received ? new Date(email.received).toLocaleString("en-IN") : "";

  const overlay = document.createElement("div");
  overlay.style.cssText = "position:fixed;inset:0;z-index:10000;background:rgba(0,0,0,.55);display:flex;align-items:center;justify-content:center;padding:20px";

  const modal = document.createElement("div");
  modal.style.cssText = "background:var(--surface);border:1px solid var(--border);border-radius:var(--r);max-width:740px;width:100%;max-height:88vh;overflow-y:auto;padding:22px;position:relative;box-shadow:0 8px 40px rgba(0,0,0,.22)";

  modal.innerHTML = `
    <div style="display:flex;align-items:start;justify-content:space-between;gap:12px;margin-bottom:14px">
      <div>
        <div style="font-size:16px;font-weight:600;color:var(--text);margin-bottom:5px">${escHtml(subject)}</div>
        <div style="font-size:12px;color:var(--muted)"><strong>From:</strong> ${escHtml(from)}</div>
        <div style="font-size:12px;color:var(--muted)"><strong>Date:</strong> ${date}</div>
      </div>
      <button class="sm" id="exp-close-btn" style="flex-shrink:0"><i class="ti ti-x"></i></button>
    </div>
    <hr style="border:none;border-top:1px solid var(--border);margin:0 0 14px">
    <div id="exp-body" style="font-size:13px;color:var(--text);white-space:pre-wrap;line-height:1.7;min-height:80px">
      <span class="ldot" style="background:var(--blue)"></span> Loading...
    </div>
    <div style="margin-top:14px;display:flex;gap:6px">
      <button class="sm primary" onclick="draftEmailReply()"><i class="ti ti-robot"></i> AI Draft Reply</button>
      <button class="sm" onclick="copyText('exp-body')"><i class="ti ti-copy"></i> Copy</button>
    </div>`;

  overlay.appendChild(modal);
  document.body.appendChild(overlay);

  modal.querySelector("#exp-close-btn").onclick = () => overlay.remove();
  overlay.onclick = (e) => { if (e.target === overlay) overlay.remove(); };

  if (email.id && creds) {
    const bodyData = await api.ewsEmailBody({ ...creds, itemId: email.id, changeKey: email.changeKey });
    const bodyText = bodyData.body || "(no content)";
    const bodyEl   = modal.querySelector("#exp-body");
    if (bodyEl) { bodyEl.textContent = bodyText; email.bodyFull = bodyText; }
  }
}

window.syncOutlookEmails = async function (btn) {
  const list = document.getElementById("ol-list");
  if (!getMsToken()) {
    showToast("Paste your Graph Explorer token in the Teams panel first.");
    return;
  }
  btn.disabled = true;
  btn.innerHTML = '<i class="ti ti-loader" style="font-size:12px;animation:spin .7s linear infinite"></i> Loading...';
  const data = await api.outlookEmails();
  btn.disabled = false;
  btn.innerHTML = '<i class="ti ti-refresh"></i> Refresh inbox';
  if (data.onPremise) {
    list.innerHTML = `<div style="font-size:12px;color:var(--muted);padding:12px;background:var(--surface2);border-radius:var(--r-sm);border:1px solid var(--border)"><strong style="color:var(--text)">On-premise Exchange detected.</strong><br>${data.error}</div>`;
    return;
  }
  if (data.error) { showToast("Outlook error: " + data.error); return; }
  const emails = data.emails || [];
  const badge  = document.getElementById("ol-badge");
  const unread = emails.filter((e) => !e.isRead).length;
  if (badge) { badge.textContent = unread; badge.style.display = unread ? "inline" : "none"; }
  if (!emails.length) { list.innerHTML = '<div style="font-size:12px;color:var(--hint);padding:16px 0;text-align:center">No emails found.</div>'; return; }
  list.innerHTML = "";
  emails.forEach((email) => {
    const row = document.createElement("div");
    row.className = "meet-row" + (email.isRead ? "" : " sel");
    row.style.cursor = "pointer";
    const from = email.from?.emailAddress?.name || email.from?.emailAddress?.address || "Unknown";
    const date = new Date(email.receivedDateTime).toLocaleDateString("en-IN", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" });
    row.innerHTML = `<div style="flex:1;min-width:0"><div class="mr-t" style="${email.isRead ? "" : "font-weight:600"}">${escHtml(email.subject || "(no subject)")}</div><div class="mr-m">${escHtml(from)} · ${date}</div><div style="font-size:11px;color:var(--hint);margin-top:2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escHtml(email.bodyPreview || "")}</div></div>`;
    row.onclick = () => previewEmail(email, row);
    list.appendChild(row);
  });
};
async function previewEmail(email, row) {
  document.querySelectorAll("#ol-list .meet-row").forEach((r) => r.classList.remove("sel"));
  row.classList.add("sel");
  _currentEmail = email;
  const det = document.getElementById("ol-detail");
  if (!det) return;
  const from = email.from?.emailAddress?.name || email.from?.emailAddress?.address || "Unknown";
  det.innerHTML = `
    <div style="border-bottom:1px solid var(--border);padding-bottom:10px;margin-bottom:10px">
      <div style="font-size:15px;font-weight:600;color:var(--text);margin-bottom:5px">${escHtml(email.subject || "(no subject)")}</div>
      <div style="font-size:12px;color:var(--muted)"><strong>From:</strong> ${escHtml(from)}</div>
      <div style="font-size:12px;color:var(--muted)"><strong>Date:</strong> ${new Date(email.receivedDateTime).toLocaleString("en-IN")}</div>
    </div>
    <div id="ol-body-m365" style="min-height:80px;margin-bottom:12px;border-radius:var(--r-sm);overflow:hidden">
      <div style="font-size:12px;color:var(--muted);padding:8px 0"><span class="ldot" style="background:var(--blue)"></span> Loading...</div>
    </div>
    <button class="primary sm" onclick="draftEmailReply()"><i class="ti ti-robot"></i> AI Draft Reply</button>`;

  if (email.id) {
    const data    = await api.outlookEmailBody(email.id);
    const content = data.body || email.bodyPreview || "(no content)";
    const isHtml  = (data.bodyType || "").toLowerCase() === "html" || /<[a-z][\s\S]*>/i.test(content);
    const bodyEl  = document.getElementById("ol-body-m365");
    if (bodyEl) {
      if (isHtml) {
        const iframe = document.createElement("iframe");
        iframe.sandbox = "allow-same-origin allow-popups";
        iframe.style.cssText = "width:100%;border:none;min-height:300px;border-radius:var(--r-sm);background:#fff";
        iframe.srcdoc = content;
        iframe.onload = () => {
          const h = iframe.contentDocument?.body?.scrollHeight;
          if (h) iframe.style.height = h + 20 + "px";
        };
        bodyEl.innerHTML = "";
        bodyEl.appendChild(iframe);
      } else {
        bodyEl.style.cssText = "font-size:13px;color:var(--text);white-space:pre-wrap;line-height:1.7;padding:8px 0";
        bodyEl.textContent = content;
      }
    }
    _currentEmail.bodyFull    = content;
    _currentEmail.bodyPreview = content.replace(/<[^>]+>/g, "").slice(0, 500);
  }
}
window.draftEmailReply = async function () {
  if (!_currentEmail) return;
  const sec  = document.getElementById("ol-draft-section");
  const txt  = document.getElementById("ol-draft-text");
  if (sec) sec.style.display = "block";
  if (txt) txt.innerHTML = `<span class="ldot" style="background:var(--blue)"></span>Drafting reply with Grok...`;
  const from = _currentEmail.from?.emailAddress?.name || _currentEmail.from?.emailAddress?.address || "";
  const data = await api.outlookDraft({
    subject:     _currentEmail.subject,
    from,
    bodyPreview: _currentEmail.bodyPreview,
  });
  if (txt) typeIn(txt, data.text || data.error || "Error drafting reply");
};
window.sendOutlookDraft = async function () {
  if (!_currentEmail) return;
  const txt = document.getElementById("ol-draft-text")?.innerText || "";
  if (!txt) { showToast("Draft is empty."); return; }
  const to = _currentEmail.from?.emailAddress?.address || _currentEmail.from?.address;
  if (!to) { showToast("Cannot determine reply-to address."); return; }
  const subject = "Re: " + (_currentEmail.subject || "");
  const creds = ewsGetCreds();
  const data = creds.ewsUrl
    ? await api.ewsSendEmail({ ...creds, to, subject, body: txt })
    : await api.outlookSend({ to, subject, body: txt });
  showToast(data.ok ? `Reply sent to ${to}!` : "Send failed: " + data.error);
};
window.sendViaTeams = function () {
  if (!_currentEmail) return;
  const txt = document.getElementById("ol-draft-text")?.innerText?.trim() || "";
  if (!txt) { showToast("Draft is empty — generate an AI reply first."); return; }
  const to = _currentEmail.from?.emailAddress?.address  // M365
           || _currentEmail.from?.address;              // EWS
  if (!to) { showToast("Cannot determine recipient email address."); return; }
  const url = `https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(to)}&message=${encodeURIComponent(txt)}`;
  window.open(url, "_blank");
};

// ── Calendar — sync from Exchange (EWS, on-premise) ───────────
window.syncOutlookCalendar = async function (btn) {
  const creds = ewsGetCreds();
  if (!creds.ewsUrl) {
    showToast("Connect to Exchange in the Teams panel first.");
    return;
  }
  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader" style="font-size:12px;animation:spin .7s linear infinite"></i> Syncing...'; }
  // Always fetch fresh — 90 days ahead so future months are populated
  const data = await api.ewsMeetings({ ...creds, days: 90 });
  if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-refresh"></i> Sync Exchange'; }
  if (data.error) { showToast("Calendar sync error: " + data.error); return; }
  _ewsMeetings = data.meetings || [];
  renderEWSMeetings(_ewsMeetings);
  updateOverviewStats(_ewsMeetings, null);
  renderOvUpcoming(_ewsMeetings);
  const statusEl = document.getElementById("cal-sync-status");
  if (statusEl) statusEl.textContent = `Last synced: ${new Date().toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" })} · ${_ewsMeetings.length} event(s)`;
  showToast(`Calendar synced — ${_ewsMeetings.length} event(s) from Exchange.`);
};

// ── Weekly report ──────────────────────────────────────────────
window.generateReport = async function () {
  const t    = document.getElementById("rp-output");
  const card = document.getElementById("rp-output-card");
  if (!t) return;
  const provider = document.getElementById("rp-provider")?.value || localStorage.getItem("ai_provider") || "groq";
  const labels = { groq: "Groq · LLaMA 3.3 70B", ollama: "Ollama · Local LLM", anthropic: "Claude · Sonnet 4.6" };
  card.style.display = "block";
  t.innerHTML = `<span class="ldot" style="background:var(--blue)"></span>Generating report with ${labels[provider] || provider}…`;
  card.scrollIntoView({ behavior: "smooth", block: "start" });
  let standupHistory = [];
  try { standupHistory = JSON.parse(localStorage.getItem("su_history") || "[]"); } catch {}
  const data = await api.generateReport({
    name:           document.getElementById("rp-name")?.value,
    dept:           document.getElementById("rp-dept")?.value,
    role:           getMyRole(),
    project:        localStorage.getItem("projectName") || "",
    startTime:      document.getElementById("rp-start")?.value,
    endTime:        document.getElementById("rp-end")?.value,
    context:        document.getElementById("rp-context")?.value,
    standupHistory,
    provider,
  });
  typeIn(t, data.text || data.error || "Error generating report");
};
window.clearReport = function () {
  const card = document.getElementById("rp-output-card");
  const t    = document.getElementById("rp-output");
  if (t) t.innerHTML = "";
  if (card) card.style.display = "none";
};
window.downloadReport = function () {
  const content = document.getElementById("rp-output")?.innerText || "";
  if (!content) return;
  const name = (document.getElementById("rp-name")?.value || "Report").replace(/\s+/g, "_");
  const fmt  = document.getElementById("rp-dl-fmt")?.value || "pdf";

  if (fmt === "doc") {
    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"><head><meta charset="utf-8"><title>Weekly Report</title></head><body style="font-family:Calibri,sans-serif;font-size:11pt;line-height:1.6;margin:2cm">${content.replace(/\n/g, "<br>")}</body></html>`;
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([html], { type: "application/msword" }));
    a.download = `Weekly_Report_${name}.doc`;
    a.click();
    URL.revokeObjectURL(a.href);
  } else {
    const win = window.open("", "_blank");
    if (!win) { showToast("Allow pop-ups to download PDF."); return; }
    win.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Weekly_Report_${name}</title><style>body{font-family:Calibri,sans-serif;font-size:11pt;line-height:1.6;margin:2cm 2.5cm;color:#1a1a1a}pre{white-space:pre-wrap;font-family:inherit}@media print{body{margin:1.5cm}}</style></head><body><pre>${content.replace(/</g,"&lt;")}</pre><script>window.onload=function(){window.print();setTimeout(function(){window.close()},500)}<\/script></html>`);
    win.document.close();
  }
};

// ── MOM generator ──────────────────────────────────────────────
// populate the MOM dropdown with this week's meetings
window.populateMOMDropdown = function () {
  const sel = document.getElementById("mm-meeting-select");
  if (!sel) return;
  const now      = new Date();
  const weekAgo  = new Date(now); weekAgo.setDate(weekAgo.getDate() - 7); weekAgo.setHours(0, 0, 0, 0);
  const weekly   = _ewsMeetings
    .filter(m => { const d = new Date(m.start); return d >= weekAgo && d <= now; })
    .sort((a, b) => new Date(b.start) - new Date(a.start));
  sel.innerHTML = weekly.length
    ? `<option value="">— select a meeting —</option>` +
      weekly.map((m) => {
        const d = new Date(m.start);
        const label = `${m.subject} — ${isNaN(d) ? m.start : d.toLocaleDateString("en-IN", { weekday:"short", day:"2-digit", month:"short" }) + " at " + d.toLocaleTimeString("en-IN", { hour:"2-digit", minute:"2-digit" })}`;
        const realIdx = _ewsMeetings.indexOf(m);
        return `<option value="${realIdx}">${label}</option>`;
      }).join("")
    : `<option value="">— no meetings found this week (sync calendar first) —</option>`;
  const info = document.getElementById("mm-meeting-info"); if (info) info.style.display = "none";
  const card = document.getElementById("mm-output-card");  if (card) card.style.display = "none";
  updateMOMGenBtn();
};
window.onMOMSelectChange = function () {
  const sel  = document.getElementById("mm-meeting-select");
  const info = document.getElementById("mm-meeting-info");
  const idx  = parseInt(sel?.value, 10);
  if (isNaN(idx) || idx < 0) { if (info) info.style.display = "none"; return; }
  const m = _ewsMeetings[idx];
  if (!m) { if (info) info.style.display = "none"; return; }
  const d = new Date(m.start);
  // Auto-fill optional detail fields
  const titleEl     = document.getElementById("mm-title-input");
  const dateEl      = document.getElementById("mm-date-input");
  const attendeesEl = document.getElementById("mm-attendees-input");
  if (titleEl)     titleEl.value     = m.subject || "";
  if (dateEl && !isNaN(d)) dateEl.value = d.toISOString().slice(0, 10);
  if (attendeesEl) attendeesEl.value = (m.attendees || []).join(", ");
  const rows = [
    `<i class="ti ti-calendar" style="font-size:11px"></i> ${isNaN(d) ? m.start : d.toLocaleDateString("en-IN", { weekday:"long", day:"2-digit", month:"long", year:"numeric" }) + " · " + d.toLocaleTimeString("en-IN", { hour:"2-digit", minute:"2-digit" })}`,
    m.dur       ? `<i class="ti ti-clock" style="font-size:11px"></i> ${m.dur}` : "",
    m.location  ? `<i class="ti ti-map-pin" style="font-size:11px"></i> ${escHtml(m.location)}` : "",
    m.attendees?.length ? `<i class="ti ti-users" style="font-size:11px"></i> ${escHtml(m.attendees.join(", "))}` : "",
  ].filter(Boolean);
  if (info) { info.innerHTML = rows.join("<br>"); info.style.display = "block"; }
};
// ── Whisper transcription ───────────────────────────────────────
let _momTranscript = "";

function updateMOMGenBtn() {
  const btn = document.getElementById("mm-gen-btn");
  if (btn) btn.disabled = !_momTranscript;
}

function showTranscriptBadge(label) {
  document.getElementById("mm-transcript-loaded-label").textContent = label;
  document.getElementById("mm-transcript-loaded").style.display = "flex";
  const ta = document.getElementById("mm-transcript-input");
  ta.value = "";
  ta.style.display = "none";
}

function showTranscriptTextarea() {
  document.getElementById("mm-transcript-loaded").style.display = "none";
  document.getElementById("mm-transcript-input").style.display = "";
}

window.clearAutoTranscript = function () {
  _momTranscript = "";
  showTranscriptTextarea();
  updateMOMGenBtn();
};

window.onMOMTranscriptChange = function () {
  _momTranscript = document.getElementById("mm-transcript-input")?.value?.trim() || "";
  updateMOMGenBtn();
};

window.onMOMTranscriptFileChosen = function (input) {
  const file = input.files[0];
  if (!file) return;
  input.value = "";

  function applyText(text) {
    showTranscriptTextarea();
    const ta = document.getElementById("mm-transcript-input");
    if (ta) ta.value = text;
    _momTranscript = text.trim();
    updateMOMGenBtn();
    showToast(`Transcript loaded from ${file.name}`);
  }

  if (file.name.endsWith(".docx")) {
    const reader = new FileReader();
    reader.onload = function (e) {
      mammoth.extractRawText({ arrayBuffer: e.target.result })
        .then(result => applyText(result.value || ""))
        .catch(() => showToast("Could not read .docx file"));
    };
    reader.readAsArrayBuffer(file);
  } else {
    const reader = new FileReader();
    reader.onload = e => applyText(e.target.result || "");
    reader.readAsText(file);
  }
};

window.onTranscribeFileChosen = function (input) {
  const file = input.files[0];
  if (!file) return;
  document.getElementById("tr-file-name").textContent = file.name;
  document.getElementById("tr-btn").disabled = false;
  document.getElementById("tr-status").textContent = "";
  _momTranscript = "";
  showTranscriptTextarea();
  updateMOMGenBtn();
};

window.transcribeRecording = async function (btn) {
  const input = document.getElementById("tr-file-input");
  if (!input?.files[0]) return;
  const file = input.files[0];
  if (file.size > 500 * 1024 * 1024) {
    document.getElementById("tr-status").textContent = "File exceeds 500 MB limit.";
    document.getElementById("tr-status").style.color = "var(--red)";
    return;
  }
  btn.disabled = true;
  const statusEl = document.getElementById("tr-status");
  statusEl.style.color = "var(--muted)";
  statusEl.textContent = file.size > 24 * 1024 * 1024
    ? "Large file detected — converting audio and transcribing… this may take a few minutes"
    : "Transcribing… this may take a minute";
  const form = new FormData();
  form.append("file", file, file.name);
  const data = await api.transcribeAudio(form);
  btn.disabled = false;
  if (data.error) {
    statusEl.style.color = "var(--red)";
    statusEl.textContent = data.error;
    return;
  }
  statusEl.textContent = "Done";
  statusEl.style.color = "var(--green)";
  const transcript = data.transcript || "";
  document.getElementById("tr-transcript").value = transcript;
  document.getElementById("tr-result").style.display = "block";
  // Auto-load into MOM generator — show badge, don't duplicate in textarea
  _momTranscript = transcript.trim();
  showTranscriptBadge("Transcript loaded from recording");
  updateMOMGenBtn();
  showToast("Transcript ready — fill in optional details and click Generate MOM");
};

window.generateMOM = async function () {
  const card = document.getElementById("mm-output-card");
  const o    = document.getElementById("mom-output");
  if (!o || !_momTranscript) return;

  // Optional detail inputs (may be auto-filled from calendar or entered manually)
  const titleInput = document.getElementById("mm-title-input")?.value?.trim();
  const dateInput  = document.getElementById("mm-date-input")?.value;
  const attsInput  = document.getElementById("mm-attendees-input")?.value?.trim();

  // Calendar meeting (optional, selected under details section)
  const sel = document.getElementById("mm-meeting-select");
  const idx = parseInt(sel?.value, 10);
  const m   = !isNaN(idx) && idx >= 0 ? _ewsMeetings[idx] : null;

  const subject   = titleInput || m?.subject || "Meeting";
  const attendees = attsInput  || (m?.attendees || []).join(", ") || "Not specified";
  const duration  = m?.dur || "";
  const context   = m?.location || "";
  let dateStr;
  if (dateInput) {
    dateStr = new Date(dateInput + "T00:00:00").toLocaleDateString("en-IN", { day:"2-digit", month:"long", year:"numeric" });
  } else if (m?.start) {
    const s = new Date(m.start);
    dateStr = isNaN(s) ? m.start : s.toLocaleDateString("en-IN", { day:"2-digit", month:"long", year:"numeric" });
  } else {
    dateStr = new Date().toLocaleDateString("en-IN", { day:"2-digit", month:"long", year:"numeric" });
  }

  card.style.display = "block";
  o.innerHTML = `<span class="ldot" style="background:var(--blue)"></span>Generating MOM…`;
  card.scrollIntoView({ behavior: "smooth", block: "start" });
  const data = await api.teamsMOM({
    subject:    subject,
    date:       dateStr,
    attendees:  attendees,
    duration:   duration,
    context:    context,
    transcript: _momTranscript,
  });
  typeIn(o, data.text || data.error || "Error generating MOM");
};
window.downloadMOMDoc = function () {
  const content = document.getElementById("mom-output")?.innerText || "";
  if (!content) return;
  const titleInput = document.getElementById("mm-title-input")?.value?.trim();
  const sel        = document.getElementById("mm-meeting-select");
  const idx        = parseInt(sel?.value, 10);
  const title      = (titleInput || _ewsMeetings[idx]?.subject || "MOM").replace(/\s+/g, "_");
  const fmt     = document.getElementById("mm-dl-fmt")?.value || "pdf";
  if (fmt === "doc") {
    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"><head><meta charset="utf-8"><title>${title}</title></head><body style="font-family:Calibri,sans-serif;font-size:11pt;line-height:1.6;margin:2cm">${content.replace(/\n/g, "<br>")}</body></html>`;
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([html], { type: "application/msword" }));
    a.download = `${title}.doc`;
    a.click();
    URL.revokeObjectURL(a.href);
  } else {
    const win = window.open("", "_blank");
    if (!win) { showToast("Allow pop-ups to download PDF."); return; }
    win.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>${title}</title><style>body{font-family:Calibri,sans-serif;font-size:11pt;line-height:1.6;margin:2cm 2.5cm;color:#1a1a1a}pre{white-space:pre-wrap;font-family:inherit}@media print{body{margin:1.5cm}}</style></head><body><pre>${content.replace(/</g,"&lt;")}</pre><script>window.onload=function(){window.print();setTimeout(function(){window.close()},500)}<\/script></html>`);
    win.document.close();
  }
};

// ── Stand-up (with auto-generate + Manager Q&A) ────────────────
let _standupTimer = null;
let _standupCtx   = null; // retained for Q&A after generation

function autoStandup() {
  clearTimeout(_standupTimer);
  const done  = document.getElementById("su-done")?.value?.trim()  || "";
  const today = document.getElementById("su-today")?.value?.trim() || "";
  if (done.length > 10 && today.length > 10) {
    _standupTimer = setTimeout(() => window.generateStandup(), 1500);
  }
}
window.generateStandup = async function () {
  const t = document.getElementById("su-output");
  if (!t) return;
  t.innerHTML = `<span class="ldot" style="background:var(--blue)"></span>Generating stand-up...`;
  const done     = document.getElementById("su-done")?.value     || "";
  const today    = document.getElementById("su-today")?.value    || "";
  const blockers = document.getElementById("su-blockers")?.value || "";
  const tone     = document.getElementById("su-format")?.value   || "casual";
  const name     = getDisplayName();
  const role     = getMyRole();
  const project  = localStorage.getItem("projectName") || "";
  const provider = localStorage.getItem("ai_provider") || "groq";
  const data = await api.generateStandup({ done, today, blockers, tone, name, role, project, provider });
  typeIn(t, data.text || data.error);
  if (data.text) {
    _standupCtx = { done, today, blockers, standupText: data.text, name, role };
    // Save to standup history for weekly report
    const entry = { date: new Date().toLocaleDateString("en-IN", { day:"2-digit", month:"short", year:"numeric" }), done, today, blockers };
    try {
      const hist = JSON.parse(localStorage.getItem("su_history") || "[]");
      hist.unshift(entry);
      localStorage.setItem("su_history", JSON.stringify(hist.slice(0, 20)));
    } catch {}
    // Save for "Load yesterday" feature
    localStorage.setItem("su_prev_done",  done);
    localStorage.setItem("su_prev_today", today);
    const qaSection = document.getElementById("su-qa-section");
    if (qaSection) {
      document.getElementById("su-qa-history").innerHTML = "";
      qaSection.style.display = "block";
      _suRenderSuggestions();
    }
  }
};
window.loadYesterday = function () {
  const doneEl  = document.getElementById("su-done");
  const todayEl = document.getElementById("su-today");
  const prevDone  = localStorage.getItem("su_prev_done")  || "";
  const prevToday = localStorage.getItem("su_prev_today") || "";
  if (doneEl  && prevToday) doneEl.value  = prevToday;
  if (todayEl && prevDone)  todayEl.value = "";
  if (!prevDone && !prevToday) showToast("No previous stand-up found — fill in the fields manually.");
};
window.standupQA = async function () {
  const input   = document.getElementById("su-qa-input");
  const history = document.getElementById("su-qa-history");
  if (!input || !history || !_standupCtx) return;
  const question = input.value.trim();
  if (!question) return;
  input.value = "";

  // Manager question bubble (left)
  const qEl = document.createElement("div");
  qEl.className = "chat-msg ai";
  qEl.innerHTML = `<div class="msg-lbl">Manager</div>${escHtml(question)}`;
  history.appendChild(qEl);

  // Answer bubble placeholder (right)
  const aEl = document.createElement("div");
  aEl.className = "chat-msg user";
  aEl.innerHTML = `<div class="msg-lbl">You</div><span class="su-qa-ans"><span class="ldot" style="background:var(--blue)"></span> Thinking...</span>`;
  history.appendChild(aEl);
  history.scrollTop = history.scrollHeight;

  const data = await api.standupQA({ ..._standupCtx, question, name: _standupCtx?.name || getDisplayName() });

  const ansEl = aEl.querySelector(".su-qa-ans");
  if (ansEl) typeIn(ansEl, data.answer || data.error);
  setTimeout(() => { history.scrollTop = history.scrollHeight; }, 100);
};
window.setSuQuestion = function (q) {
  const inp = document.getElementById("su-qa-input");
  if (inp) { inp.value = q; inp.focus(); }
};
function _suRenderSuggestions() {
  const wrap = document.getElementById("su-qa-suggestions");
  if (!wrap) return;
  const suggestions = [
    "What's your ETA for today's tasks?",
    "Any risk of missing today's deadline?",
    "How long has the blocker been pending?",
    "What's the business impact of the blocker?",
    "Can you walk me through what you completed yesterday?",
    "Do you need help to unblock yourself?",
  ];
  wrap.innerHTML = suggestions
    .map((s) => `<button class="sm" style="font-size:11px;color:var(--muted)" onclick="setSuQuestion(${JSON.stringify(s)})">${escHtml(s)}</button>`)
    .join("");
}

// ── Leave & Resources ──────────────────────────────────────────
// ── Team members (leave register) ──────────────────────────────
const LV_KEY = "teamMembers";
let _lvCache = [];
function lvLoadMembers() { return _lvCache.length ? _lvCache : (()=>{ try { _lvCache = JSON.parse(localStorage.getItem(LV_KEY) || "[]"); return _lvCache; } catch { return []; } })(); }
function lvSaveMembers(arr) { _lvCache = arr; localStorage.setItem(LV_KEY, JSON.stringify(arr)); }

async function lvFetchFromServer() {
  try {
    const data = await api.teamMembers();
    if (data && Array.isArray(data.members)) {
      _lvCache = data.members;
      localStorage.setItem(LV_KEY, JSON.stringify(_lvCache));
    }
  } catch {}
}

let _lvPollTimer = null;
function lvStartPolling() {
  if (_lvPollTimer) return;
  _lvPollTimer = setInterval(async () => {
    await lvFetchFromServer();
    refreshLeaveCards();
  }, 30000);
}
function lvStopPolling() {
  if (_lvPollTimer) { clearInterval(_lvPollTimer); _lvPollTimer = null; }
}

function computeCurrentStatus(member) {
  // If member has leave dates, check if today falls within range
  if (member.leaveFrom && member.leaveTo) {
    const today = new Date(); today.setHours(0,0,0,0);
    const from  = new Date(member.leaveFrom);
    const to    = new Date(member.leaveTo);
    if (today >= from && today <= to) return "on-leave";
    if (today > to) return "available";
  }
  return member.status || "available";
}

function autoComputeStatuses() {
  const members = lvLoadMembers();
  let changed = false;
  members.forEach(m => {
    const computed = computeCurrentStatus(m);
    if (computed !== m.status) { m.status = computed; changed = true; }
  });
  if (changed) lvSaveMembers(members);
}

function refreshLeaveCards() {
  autoComputeStatuses();
  renderStakeholders();
  populateMemberSelect();
}

function populateMemberSelect() {
  const sel = document.getElementById("lv-member-select");
  if (!sel) return;
  const members = lvLoadMembers();
  const cur = sel.value;
  sel.innerHTML = '<option value="">— select team member —</option>' +
    members.map((m, i) => `<option value="${i}">${escHtml(m.name)} (${escHtml(m.role)})</option>`).join("");
  if (cur) sel.value = cur;
}

function renderStakeholders(preloaded) {
  const tbody  = document.getElementById("stakeholder-tbody");
  const empty  = document.getElementById("lv-empty-row");
  if (!tbody) return;
  const members = preloaded || lvLoadMembers();
  tbody.querySelectorAll("tr.lv-member-row").forEach(r => r.remove());
  if (empty) empty.style.display = members.length ? "none" : "";
  members.forEach((m, i) => {
    const onLeave = m.status === "on-leave";
    const badge = onLeave
      ? '<span class="badge b-amber">On Leave</span>'
      : '<span class="badge b-green">Available</span>';
    const tr = document.createElement("tr");
    tr.className = "lv-member-row";
    tr.dataset.idx = i;
    tr.innerHTML = '<td>' + (i + 1) + '</td><td style="font-weight:500">' + escHtml(m.name) + '</td><td>' + escHtml(m.role) + '</td><td>' + badge + '</td><td><button class="sm lv-menu-btn" data-mi="' + i + '" onclick="lvToggleMenu(this,' + i + ')" style="padding:2px 10px;font-size:18px;line-height:1;letter-spacing:2px;border-radius:6px;min-width:34px">⋯</button></td>';
    tbody.appendChild(tr);
  });
  const availCount  = members.filter(m => m.status !== "on-leave").length;
  const leaveCount  = members.filter(m => m.status === "on-leave").length;
  const leaveNames  = members.filter(m => m.status === "on-leave").map(m => m.name).join(", ") || "—";
  const availEl     = document.getElementById("lv-avail-count");
  const totalEl     = document.getElementById("lv-total-sub");
  const leaveEl     = document.getElementById("lv-leave-count");
  const leaveNamesEl = document.getElementById("lv-leave-names");
  if (availEl)     availEl.textContent = availCount;
  if (totalEl)     totalEl.textContent = `of ${members.length} member${members.length !== 1 ? "s" : ""}`;
  if (leaveEl)     leaveEl.textContent = leaveCount;
  if (leaveNamesEl) leaveNamesEl.textContent = leaveNames;
}

window.addTeamMember = function () {
  const name = document.getElementById("lv-new-name")?.value?.trim();
  const role = document.getElementById("lv-new-role")?.value?.trim();
  if (!name) { showToast("Enter a name."); return; }
  if (!role) { showToast("Enter a role."); return; }
  const members = lvLoadMembers();
  members.push({ name, role, status: "available" });
  lvSaveMembers(members);
  document.getElementById("lv-new-name").value = "";
  document.getElementById("lv-new-role").value = "";
  document.getElementById("lv-new-name").focus();
  refreshLeaveCards();
  showToast(`Team member "${name}" added successfully.`);
};

window.renderStakeholders = renderStakeholders;

window.removeTeamMember = function (i) {
  const members = lvLoadMembers();
  const removed = members[i]?.name || "Member";
  members.splice(i, 1);
  lvSaveMembers(members);
  lvCloseMenu();
  refreshLeaveCards();
  showToast('"' + removed + '" removed from team.');
};

// ── Singleton action menu (fixed-position, avoids overflow:hidden clipping) ──
let _lvMenuIdx = -1;
function lvGetMenuEl() {
  let el = document.getElementById("lv-action-menu");
  if (!el) {
    el = document.createElement("div");
    el.id = "lv-action-menu";
    document.body.appendChild(el);
  }
  el.style.cssText = "display:none;position:fixed;z-index:9999;background:var(--surface);border:1px solid var(--border);border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.2);min-width:140px;overflow:hidden";
  return el;
}
function lvCloseMenu() {
  const el = document.getElementById("lv-action-menu");
  if (el) el.style.display = "none";
  _lvMenuIdx = -1;
}

window.lvToggleMenu = function (btn, i) {
  const existing = document.getElementById("lv-action-menu");
  if (_lvMenuIdx === i && existing && existing.style.display !== "none") { lvCloseMenu(); return; }
  const menu = lvGetMenuEl();
  _lvMenuIdx = i;
  const btnStyle = "width:100%;text-align:left;padding:8px 14px;background:none;border:none;font-size:13px;cursor:pointer;display:flex;align-items:center;gap:8px;";
  menu.innerHTML =
    '<button onclick="editTeamMember(' + i + ')" style="' + btnStyle + 'color:var(--text)"><i class="ti ti-pencil"></i> Edit</button>'
    + '<div style="height:1px;background:var(--border);margin:2px 0"></div>'
    + '<button onclick="removeTeamMember(' + i + ')" style="' + btnStyle + 'color:var(--red)"><i class="ti ti-trash"></i> Delete</button>';
  const rect = btn.getBoundingClientRect();
  menu.style.display = "block";
  const mw = menu.offsetWidth;
  let left = rect.right - mw;
  if (left < 8) left = 8;
  menu.style.top  = (rect.bottom + 4) + "px";
  menu.style.left = left + "px";
  menu.style.right = "auto";
};

window.editTeamMember = function (i) {
  lvCloseMenu();
  const members = lvLoadMembers();
  const m = members[i];
  if (!m) return;
  const tr = document.querySelector("tr.lv-member-row[data-idx='" + i + "']");
  if (!tr) return;
  tr.innerHTML =
    '<td>' + (i + 1) + '</td>'
    + '<td><input id="lv-edit-name-' + i + '" value="' + escHtml(m.name) + '" style="margin:0;padding:4px 8px;font-size:12px;width:100%"></td>'
    + '<td><input id="lv-edit-role-' + i + '" value="' + escHtml(m.role) + '" style="margin:0;padding:4px 8px;font-size:12px;width:100%"></td>'
    + '<td><select id="lv-edit-status-' + i + '" style="margin:0;padding:4px 8px;font-size:12px;width:100%"><option value="available"' + (m.status !== "on-leave" ? " selected" : "") + '>Available</option><option value="on-leave"' + (m.status === "on-leave" ? " selected" : "") + '>On Leave</option></select></td>'
    + '<td style="display:flex;gap:4px;padding:6px 4px"><button class="sm primary" onclick="saveTeamMember(' + i + ')" style="padding:3px 10px;font-size:12px">Save</button><button class="sm" onclick="window.renderStakeholders()" style="padding:3px 8px;font-size:12px">Cancel</button></td>';
};

window.saveTeamMember = function (i) {
  const name   = document.getElementById("lv-edit-name-" + i)?.value?.trim();
  const role   = document.getElementById("lv-edit-role-" + i)?.value?.trim();
  const status = document.getElementById("lv-edit-status-" + i)?.value || "available";
  if (!name) { showToast("Name cannot be empty."); return; }
  if (!role)  { showToast("Role cannot be empty."); return; }
  const members = lvLoadMembers();
  if (!members[i]) return;
  members[i] = { ...members[i], name, role, status };
  lvSaveMembers(members);
  refreshLeaveCards();
  showToast(name + " updated.");
};

document.addEventListener("click", function (e) {
  if (!e.target.closest(".lv-menu-btn") && !e.target.closest("#lv-action-menu")) {
    lvCloseMenu();
  }
});

// set today as default leave dates and load saved emails
(function initLeaveForm() {
  const today = new Date().toISOString().slice(0, 10);
  const from  = document.getElementById("lv-from");
  const to    = document.getElementById("lv-to");
  if (from && !from.value) from.value = today;
  if (to   && !to.value)   to.value   = today;
  // restore saved email settings
  const mgr = localStorage.getItem("lv_mgr_email");
  const hr  = localStorage.getItem("lv_hr_email");
  if (mgr) { const el = document.getElementById("lv-to-email"); if (el) el.value = mgr; }
  if (hr)  { const el = document.getElementById("lv-cc-email");  if (el) el.value = hr; }
  // persist email changes
  document.getElementById("lv-to-email")?.addEventListener("change", e => localStorage.setItem("lv_mgr_email", e.target.value));
  document.getElementById("lv-cc-email")?.addEventListener("change", e => localStorage.setItem("lv_hr_email",  e.target.value));
  lvFetchFromServer().then(() => refreshLeaveCards());
})();

window.syncLeaveFromExchange = async function (btn) {
  const creds = ewsGetCreds();
  if (!creds.ewsUrl) { showToast("Connect to Exchange in Teams panel first."); return; }
  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader" style="font-size:12px;animation:spin .7s linear infinite"></i> Syncing...'; }

  const data = await api.ewsMeetings({ ...creds, days: 14 });

  if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-refresh"></i> Sync from Exchange'; }

  const status = document.getElementById("lv-sync-status");
  if (data.error) {
    if (status) { status.style.display = "block"; status.textContent = "Sync error: " + data.error; }
    return;
  }

  const meetings = data.meetings || [];
  const today    = new Date(); today.setHours(0,0,0,0);
  const twoWeeks = new Date(today.getTime() + 14 * 86400000);

  // Detect OOF/all-day leave-type events for self
  const leaveEvents = meetings.filter((m) => {
    const s = new Date(m.start);
    return s >= today && s <= twoWeeks && (m.response === "Organizer" || m.isOnline === false);
  });

  const plannedCount = document.getElementById("lv-planned-count");
  if (plannedCount) plannedCount.textContent = leaveEvents.length;

  if (status) {
    status.style.display = "block";
    status.innerHTML     = leaveEvents.length
      ? `${leaveEvents.length} upcoming event(s) in next 14 days synced from Exchange calendar.`
      : "No upcoming leave detected in next 14 days.";
  }

  showToast(`Leave sync complete — ${meetings.length} events loaded.`);
};

window.applyLeave = async function () {
  const fromDate   = document.getElementById("lv-from")?.value || "";
  const toDate     = document.getElementById("lv-to")?.value   || "";
  const leaveType  = document.getElementById("lv-type")?.value || "Planned leave";
  const senderName = getDisplayName();

  if (!fromDate) { showToast("Select a leave start date."); return; }
  if (!senderName) { showToast("Enter your name in the Weekly Report name field first."); return; }

  // Mark the person on leave in the team members list
  const members = lvLoadMembers();
  const idx = members.findIndex(mb => mb.name.toLowerCase() === senderName.toLowerCase());
  if (idx !== -1) {
    members[idx].status    = "on-leave";
    members[idx].leaveFrom = fromDate;
    members[idx].leaveTo   = toDate || fromDate;
  } else {
    // Auto-add self if not in list
    members.push({ name: senderName, role: localStorage.getItem("rp_role") || "Team Member", status: "on-leave", leaveFrom: fromDate, leaveTo: toDate || fromDate });
  }
  lvSaveMembers(members);

  // Persist to server
  try { await api.saveTeamMembers(members); } catch {}

  refreshLeaveCards();

  const fmt = (d) => d ? new Date(d).toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" }) : "";
  const range = fromDate === toDate ? fmt(fromDate) : `${fmt(fromDate)} – ${fmt(toDate)}`;
  showToast(`Leave applied for ${senderName}: ${range}`);
};

window.generateLeaveEmail = async function () {
  const t        = document.getElementById("lv-output");
  const controls = document.getElementById("lv-email-controls");
  if (!t) return;
  t.style.display    = "block";
  t.textContent      = "";
  t.setAttribute("contenteditable", "false");
  t.innerHTML        = `<span class="ldot" style="background:var(--blue)"></span> Generating leave email…`;
  if (controls) controls.style.display = "none";

  const leaveType = document.getElementById("lv-type")?.value || "Planned leave";
  const fromDate  = document.getElementById("lv-from")?.value || "";
  const toDate    = document.getElementById("lv-to")?.value || "";
  const reason    = document.getElementById("lv-reason")?.value?.trim() || "";
  const toEmail   = document.getElementById("lv-to-email")?.value?.trim() || "";
  const fmt = (d) => d ? new Date(d).toLocaleDateString("en-IN", { day: "2-digit", month: "long", year: "numeric" }) : d;
  const dateRange = fromDate === toDate ? fmt(fromDate) : `${fmt(fromDate)} to ${fmt(toDate)}`;
  const senderName = getDisplayName() || "Team Member";
  const project    = localStorage.getItem("projectName") || "";
  const projectCtx = project ? ` Currently working on the ${project} project.` : "";
  const prompt = `Write a formal, professional ${leaveType.toLowerCase()} request email from ${senderName} (ESDS Software Solution Pvt. Ltd.)` +
    ` for dates: ${dateRange}.` +
    (reason ? ` Reason: ${reason}.` : "") +
    ` To: ${toEmail || "Manager"}.${projectCtx}` +
    ` Include: polite greeting, specific leave dates, brief reason, assurance of work handover, request for approval, and a thank-you close. Do not include a subject line. Use a professional letter format with proper paragraph spacing.`;
  const data = await api.chat(prompt, "You write formal, courteous workplace emails in proper letter format.", null, getWorkContext());
  const text = data.text || data.error || "";
  t.setAttribute("contenteditable", "true");
  typeIn(t, text);
  if (controls) controls.style.display = "block";
};

window.clearLeaveEmail = function () {
  const t        = document.getElementById("lv-output");
  const controls = document.getElementById("lv-email-controls");
  if (t) { t.textContent = ""; t.style.display = "none"; }
  if (controls) controls.style.display = "none";
};

window.sendLeaveEmail = async function (btn) {
  const emailText = document.getElementById("lv-output")?.innerText?.trim();
  if (!emailText || emailText.includes("Generating leave email")) {
    showToast("Generate the email first, then send.");
    return;
  }
  const toEmail  = document.getElementById("lv-to-email")?.value?.trim();
  const ccEmail  = document.getElementById("lv-cc-email")?.value?.trim();
  if (!toEmail) { showToast("Enter a recipient email (To field)."); return; }

  const creds = ewsGetCreds();
  if (!creds.ewsUrl) { showToast("Connect to Exchange in the Teams panel first."); return; }

  const leaveType = document.getElementById("lv-type")?.value || "Leave";
  const fromDate  = document.getElementById("lv-from")?.value || "";
  const toDate    = document.getElementById("lv-to")?.value || "";
  const fmt = (d) => d ? new Date(d).toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" }) : d;
  const subject   = fromDate === toDate
    ? `${leaveType} Request — ${fmt(fromDate)}`
    : `${leaveType} Request — ${fmt(fromDate)} to ${fmt(toDate)}`;

  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader" style="font-size:12px;animation:spin .7s linear infinite"></i> Sending...'; }

  const data = await api.ewsSendEmail({ ...creds, to: toEmail, cc: ccEmail, subject, body: emailText });

  if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-send"></i> Send via Outlook'; }

  if (data.ok) {
    showToast("Leave email sent via Outlook!");
    // Auto-mark matching team member as On Leave
    const senderName = document.getElementById("rp-name")?.value?.trim() || "";
    if (senderName) {
      const members = lvLoadMembers();
      const idx = members.findIndex(function (mb) { return mb.name.toLowerCase() === senderName.toLowerCase(); });
      if (idx !== -1 && members[idx].status !== "on-leave") {
        members[idx].status = "on-leave";
        lvSaveMembers(members);
        renderStakeholders();
        showToast(members[idx].name + " marked as On Leave in Team Members.");
      }
    }
    // Add leave to calendar
    if (fromDate && toDate) {
      const start = new Date(fromDate);
      const end   = new Date(toDate);
      for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
        const k = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
        if (!events[k]) events[k] = [];
        if (!events[k].some((e) => e.title === leaveType)) {
          events[k].push({ title: leaveType, time: "09:00 AM", dur: "All day", att: "", type: "block", h: 9, m: 0 });
        }
      }
      renderCalendar(events);
      renderOvMeetings(events);
      showToast("Leave added to calendar.");
    }
  } else {
    showToast("Send failed: " + (data.error || "Unknown error"));
  }
};

// ── General assistant ──────────────────────────────────────────

function renderMarkdown(text) {
  let html = escHtml(text);
  // Code blocks (```...```)
  html = html.replace(/```([\s\S]*?)```/g, (_, code) =>
    `<pre style="background:var(--surface2);border-radius:5px;padding:9px 12px;overflow-x:auto;font-family:var(--mono);font-size:12px;margin:8px 0;white-space:pre-wrap">${code.trim()}</pre>`);
  // Inline code
  html = html.replace(/`([^`]+)`/g,
    '<code style="background:var(--surface2);padding:1px 5px;border-radius:3px;font-family:var(--mono);font-size:12px">$1</code>');
  // Bold
  html = html.replace(/\*\*([^*\n]+)\*\*/g, '<strong>$1</strong>');
  // H3 / H2 / H1
  html = html.replace(/^### (.+)$/gm, '<div style="font-size:13px;font-weight:600;color:var(--text);margin:10px 0 3px">$1</div>');
  html = html.replace(/^## (.+)$/gm,  '<div style="font-size:14px;font-weight:600;color:var(--text);margin:12px 0 4px;border-bottom:1px solid var(--border);padding-bottom:3px">$1</div>');
  html = html.replace(/^# (.+)$/gm,   '<div style="font-size:15px;font-weight:600;color:var(--text);margin:14px 0 5px">$1</div>');
  // Numbered list items
  html = html.replace(/^\d+\.\s+(.+)$/gm, (m) => {
    const [, num, rest] = m.match(/^(\d+)\.\s+(.+)$/);
    return `<div style="display:flex;gap:7px;margin:2px 0;align-items:baseline"><span style="color:var(--blue);font-weight:600;flex-shrink:0;min-width:18px">${num}.</span><span>${rest}</span></div>`;
  });
  // Bullet list items (- or *)
  html = html.replace(/^[-*]\s+(.+)$/gm,
    '<div style="display:flex;gap:7px;margin:2px 0;align-items:baseline"><span style="color:var(--blue);flex-shrink:0">•</span><span>$1</span></div>');
  // Horizontal rule
  html = html.replace(/^---+$/gm, '<hr style="border:none;border-top:1px solid var(--border);margin:10px 0">');
  // Paragraph breaks
  html = html.replace(/\n\n/g, '<div style="height:7px"></div>');
  html = html.replace(/\n/g, '<br>');
  return html;
}

// ── User context helpers ───────────────────────────────────────
function getDisplayName() {
  return document.getElementById("rp-name")?.value?.trim()
    || localStorage.getItem("rp_name")
    || localStorage.getItem("firstName")
    || "";
}

function getMyRole() {
  return document.getElementById("rp-dept")?.value?.trim()
    || localStorage.getItem("rp_dept")
    || "";
}

function buildChatGreeting() {
  const name    = getDisplayName();
  const project = localStorage.getItem("projectName") || "";
  const hi      = name ? `Hi ${name.split(" ")[0]}!` : "Hi!";
  const proj = project ? ` How can I help you with **${project}**?` : " How can I help you today?";
  return `${hi}${proj}`;
}

function getWorkContext() {
  return {
    userName:    getDisplayName(),
    userRole:    getMyRole(),
    recentWork:  localStorage.getItem("su_prev_done")  || "",
    currentPlan: localStorage.getItem("su_prev_today") || "",
  };
}

function getChatSuggestions() {
  const project = localStorage.getItem("projectName") || "";
  const p = project ? project : "the project";
  return [
    `Write user stories for ${p}`,
    `Create acceptance criteria for ${p}`,
    `Create a BRD outline for ${p}`,
    `Identify risks in ${p}`,
    `Create an RTM template for ${p}`,
    `Write a gap analysis for ${p}`,
    `Draft a formal email regarding ${p}`,
    "Key differences between BRD and FRD",
    "Write a change request for",
    "Create a RACI matrix for",
    "Write test cases for",
    "Summarize key requirements for",
    "Write an executive summary for",
    "Define KPIs for the team",
    "Draft meeting agenda for project review",
    "Create stakeholder register for",
    "Write a project charter for",
    "How do I write a good use case?",
    "What is the difference between functional and non-functional requirements?",
    "Explain MoSCoW prioritization",
    "How should I structure a requirements traceability matrix?",
  ];
}

function appendFollowupChips(msgEl, lastMsg) {
  const project = localStorage.getItem("projectName") || "";
  const msg = lastMsg.toLowerCase();
  let chips = [];
  if (/brd|requirement|user stor/i.test(msg))
    chips = ["Create acceptance criteria", "Write an RTM", "Identify risks", "Add non-functional requirements"];
  else if (/risk|issue|blocker/i.test(msg))
    chips = ["Create a mitigation plan", "Write a risk register", "Prioritize these risks", "Draft a status update"];
  else if (/email|letter|draft/i.test(msg))
    chips = ["Make it more formal", "Make it shorter", "Add a follow-up reminder", "Translate to Hindi"];
  else if (/test|qa|bug/i.test(msg))
    chips = ["Write automation test cases", "Create a test plan", "Prioritize test scenarios", "Write a bug report"];
  else
    chips = [`Write user stories for ${project || "the project"}`, "Summarize this", "Create a checklist", "Draft an email about this"];

  if (!chips.length) return;
  const chipRow = document.createElement("div");
  chipRow.className = "chat-followups";
  chipRow.innerHTML = chips.map(c =>
    `<button class="chat-followup-chip" onclick="document.getElementById('chat-input').value=${JSON.stringify(c)};sendChatMessage()">${escHtml(c)}</button>`
  ).join("");
  msgEl.appendChild(chipRow);
}

window.chatAutocomplete = function (val) {
  const box = document.getElementById("chat-suggestions");
  if (!box) return;
  const q = val.trim().toLowerCase();
  if (q.length < 2) { box.style.display = "none"; return; }
  const matches = getChatSuggestions().filter((s) => s.toLowerCase().includes(q)).slice(0, 6);
  if (!matches.length) { box.style.display = "none"; return; }
  box.style.display = "block";
  box.innerHTML = matches.map((s) =>
    `<div class="chat-suggest-item" onclick="pickSuggestion(${JSON.stringify(s)})">${escHtml(s)}</div>`
  ).join("");
};

window.pickSuggestion = function (text) {
  const inp = document.getElementById("chat-input");
  const box = document.getElementById("chat-suggestions");
  if (inp) inp.value = text;
  if (box) box.style.display = "none";
  if (inp) inp.focus();
};

window.sendChatMessage = async function () {
  const inp = document.getElementById("chat-input");
  const box = document.getElementById("chat-suggestions");
  const msg = inp?.value?.trim();
  if (!msg) return;
  if (inp) inp.value = "";
  if (box) box.style.display = "none";
  const msgs = document.getElementById("chat-messages");
  if (!msgs) return;
  msgs.innerHTML += `<div class="chat-msg user"><div class="msg-lbl">You</div>${escHtml(msg)}</div>`;
  const thinking = document.createElement("div");
  thinking.className = "chat-msg ai";
  thinking.innerHTML = `<div class="msg-lbl">Assistant</div><span class="ldot" style="background:var(--muted)"></span> Thinking…`;
  msgs.appendChild(thinking);
  msgs.scrollTop = msgs.scrollHeight;
  const data = await api.chat(msg, null, null, getWorkContext());
  if (data.error) {
    thinking.innerHTML =
      `<div class="msg-lbl" style="color:var(--red)">Error</div>` +
      `<div style="font-size:13px;color:var(--red);margin-bottom:8px">${escHtml(data.error)}</div>` +
      `<div style="font-size:12px;color:var(--muted)">Switch to a different provider: ` +
      `<button class="sm" style="font-size:11px" onclick="setAIProvider('groq')">⚡ Groq</button> ` +
      `<button class="sm" style="font-size:11px;margin-left:4px" onclick="setAIProvider('anthropic')">🧠 Claude</button></div>`;
  } else {
    thinking.innerHTML = `<div class="msg-lbl">Assistant</div>${renderMarkdown(data.text || "No response")}`;
    appendFollowupChips(thinking, msg);
  }
  msgs.scrollTop = msgs.scrollHeight;
};

window.quickChat = function (msg) {
  const inp = document.getElementById("chat-input");
  if (inp) inp.value = msg;
  window.sendChatMessage();
  const aiNav = document.querySelector('[data-panel="ai"]');
  if (aiNav) window.nav(aiNav);
};

window.clearChat = function () {
  const msgs = document.getElementById("chat-messages");
  if (msgs) msgs.innerHTML = `<div class="chat-msg ai"><div class="msg-lbl">Assistant</div>${escHtml(buildChatGreeting())}</div>`;
};

window.setAIProvider = function (provider) {
  localStorage.setItem("ai_provider", provider);
  // Sync all provider selectors across panels
  ["ai-provider-select", "rp-provider", "ov-ai-platform"].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.value = provider;
  });
  const labels = { groq: "Groq · LLaMA 3.3 70B", ollama: "Ollama · Local LLM", anthropic: "Claude · Sonnet 4.6" };
  showToast(`AI model: ${labels[provider] || provider}`);
};

window.updateProjectName = function (val) {
  localStorage.setItem("projectName", val || "");
  const el = document.getElementById("sb-project-name");
  if (el) { el.textContent = val || ""; el.style.display = val?.trim() ? "block" : "none"; }
};

window.submitProjectSetup = function () {
  const val      = document.getElementById("ov-project")?.value?.trim() || "";
  const provider = document.getElementById("ov-ai-platform")?.value || "groq";
  const errEl    = document.getElementById("ov-project-err");

  if (!val) {
    if (errEl) { errEl.textContent = "Project name is required."; errEl.style.display = "block"; }
    const inp = document.getElementById("ov-project");
    if (inp) { inp.style.border = "1.5px solid var(--red)"; inp.focus(); }
    return;
  }
  if (errEl) errEl.style.display = "none";
  const inp = document.getElementById("ov-project");
  if (inp) inp.style.border = "";

  window.updateProjectName(val);
  window.setAIProvider(provider);
  applyNavLock();
  // Update chat greeting with new project name
  const msgs = document.getElementById("chat-messages");
  if (msgs && msgs.children.length <= 1) {
    msgs.innerHTML = `<div class="chat-msg ai"><div class="msg-lbl">Assistant</div>${escHtml(buildChatGreeting())}</div>`;
  }
  showOvMain();
  showToast(`Project "${val}" saved — AI: ${provider}`);
};

// set today as default for the calendar Add Event date field
(function() { const el = document.getElementById("ev-date"); if (el && !el.value) el.value = new Date().toISOString().slice(0, 10); })();

// ── Populate Weekly Report reporter fields from login ──────────
(function initReportFields() {
  const nameEl = document.getElementById("rp-name");
  const deptEl = document.getElementById("rp-dept");
  if (nameEl) {
    // Prefer saved full name, fall back to firstName from EWS login
    const savedName = localStorage.getItem("rp_name") || "";
    const firstName = localStorage.getItem("firstName") || "";
    const ewsUser   = localStorage.getItem("ewsUsername") || "";
    // Build display name from EWS username if no saved name: "domain\mahesh.beesu" → "Mahesh Beesu"
    let derivedName = "";
    if (ewsUser) {
      const u = ewsUser.includes("\\") ? ewsUser.split("\\")[1] : ewsUser;
      derivedName = u.split(/[.\-_]/).map(p => p.charAt(0).toUpperCase() + p.slice(1).toLowerCase()).join(" ");
    }
    nameEl.value = savedName || derivedName || firstName;
    nameEl.addEventListener("change", () => localStorage.setItem("rp_name", nameEl.value));
  }
  if (deptEl) {
    deptEl.value = localStorage.getItem("rp_dept") || "";
    deptEl.addEventListener("change", () => localStorage.setItem("rp_dept", deptEl.value));
  }
})();

// ── Restore project name + AI provider on load ─────────────────
(function initProjectSetup() {
  const project  = localStorage.getItem("projectName") || "";
  const provider = localStorage.getItem("ai_provider") || "groq";
  const projInp  = document.getElementById("ov-project");
  const aiSel    = document.getElementById("ov-ai-platform");
  if (projInp) projInp.value = project;
  if (aiSel)   aiSel.value   = provider;
  if (project) { window.updateProjectName(project); showOvMain(); }
})();

// ── Copy / share helpers (used inline from HTML) ───────────────
window.copyText = copyText;

// ── OAuth token from URL ───────────────────────────────────────
function handleOAuthReturn() {
  const params = new URLSearchParams(window.location.search);
  const tok    = params.get("sp_token");
  const rt     = params.get("sp_refresh");
  const spErr  = params.get("sp_error");
  if (tok) {
    localStorage.setItem("spToken", tok);
    if (rt) localStorage.setItem("spRefreshToken", rt);
    const el = document.getElementById("sp-token");
    if (el) el.value = tok;
    history.replaceState({}, "", window.location.pathname);
    scheduleTokenRefreshTimer(tok);
    showToast("Microsoft 365 connected successfully.");
    updateTokenUI();
  }
  if (spErr) {
    showToast("OAuth error: " + decodeURIComponent(spErr));
    history.replaceState({}, "", window.location.pathname);
  }
}

// ── Init ───────────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {
  // Nav lock
  applyNavLock();

  // Chat greeting
  const chatMsgs = document.getElementById("chat-messages");
  if (chatMsgs && chatMsgs.children.length <= 1) {
    chatMsgs.innerHTML = `<div class="chat-msg ai"><div class="msg-lbl">Assistant</div>${escHtml(buildChatGreeting())}</div>`;
  }

  // Calendar
  renderCalendar(events);
  renderSchedule(todayKey(), events);
  renderOvMeetings(events);
  const badgeEl = document.getElementById("cal-badge");
  if (badgeEl) badgeEl.textContent = (events[todayKey()] || []).length;

  // Restore saved AI provider
  const savedProvider = localStorage.getItem("ai_provider") || "groq";
  ["ai-provider-select", "rp-provider"].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.value = savedProvider;
  });

  // Charts (defer so Canvas is ready)
  setTimeout(() => {
    initDocChart();
    updateDocChart();
    updateOverviewStats([], null);
  }, 300);

  // Documents — register chart refresh hook and load
  window._refreshDocStats = () => { updateDocChart(); updateOverviewStats(_ewsMeetings, null); };
  filterDocs();

  // Restore saved keys
  const ak = localStorage.getItem("anthropicKey");
  const spTok = localStorage.getItem("spToken");
  if (ak     && document.getElementById("cfg-ak"))   document.getElementById("cfg-ak").value   = ak;
  if (spTok  && document.getElementById("sp-token")) document.getElementById("sp-token").value = spTok;

  // Handle OAuth return (stores refresh token, schedules timer)
  handleOAuthReturn();

  // Update Microsoft 365 token UI + check stored token expiry
  updateTokenUI();
  const storedTok = localStorage.getItem("spToken");
  if (storedTok) {
    const expiry = parseJwtExpiry(storedTok);
    if (expiry && expiry <= new Date()) {
      // Already expired at startup — try silent refresh or clear
      handleTokenExpired();
    } else if (expiry) {
      scheduleTokenRefreshTimer(storedTok);
    }
  }

  // Stand-up auto-generate on input
  document.getElementById("su-done")?.addEventListener("input",  autoStandup);
  document.getElementById("su-today")?.addEventListener("input", autoStandup);

  // Restore Exchange state and navigate based on stored credentials
  const ewsCreds = ewsGetCreds();
  const { ewsUrl, username, password } = ewsCreds;
  ewsRestoreUI();
  window.updateWelcomeMessage();
  if (ewsUrl && username && password) {
    // Already connected — remove login-mode, show sidebar, go to Dashboard
    document.getElementById("app")?.classList.remove("login-mode");
    stopTmBackground();
    navTo("ov");
    // Auto-load meetings in background
    api.ewsMeetings({ ewsUrl, username, password }).then((data) => {
      if (data.error) return;
      _ewsMeetings = data.meetings || [];
      renderEWSMeetings(_ewsMeetings);
      updateOverviewStats(_ewsMeetings, null);
      renderOvUpcoming(_ewsMeetings);
    });
  } else {
    // Not connected — start the AI background animation
    startTmBackground();
  }
});
