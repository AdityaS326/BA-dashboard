// frontend/src/main.js
// Main entry point — imports modules and wires up all event handlers.

import { api }                  from "./utils/api.js";
import { typeIn, copyText, showToast, escHtml } from "./utils/ui.js";
import { seedEvents, todayKey } from "./utils/calendar.js";
import { renderOvMeetings, initHealthChart, initDocChart, filterHealthChart } from "./pages/overview.js";
import { renderCalendar, renderSchedule, calNav as _calNav, addEvent as _addEvent, checkAndShowReminder } from "./pages/calendar.js";
import { DOCS, renderDocs, filterDocs, addNewDoc } from "./pages/documents.js";

// ── State ──────────────────────────────────────────────────────
let events = seedEvents();
let healthChart = null;
let _tokenTimer      = null;   // proactive expiry warning timer
let _tokenRefreshing = false;  // prevent concurrent refresh calls

// ── Navigation ─────────────────────────────────────────────────
const PAGE_TITLES = {
  ov: "Dashboard",    cal: "Calendar & meetings",
  tm: "Teams meetings", ol: "Outlook inbox",
  wr: "Weekly report",  mm: "MOM generator",
  su: "Stand-up generator", dc: "Document repository",
  lv: "Leave & resources",  ai: "General assistant",
};
window.nav = function (el) {
  document.querySelectorAll(".panel").forEach((p) => p.classList.remove("active"));
  document.querySelectorAll(".nav-item").forEach((n) => n.classList.remove("active"));
  const id    = el.dataset.panel;
  const panel = document.getElementById("p-" + id);
  if (panel) panel.classList.add("active");
  el.classList.add("active");
  const titleEl = document.getElementById("page-title");
  if (titleEl) titleEl.textContent = PAGE_TITLES[id] || id;
  if (id === "cal") {
    renderCalendar(events);
    renderSchedule(todayKey(), events);
    // Auto-pull Exchange calendar if connected and not yet loaded
    const creds = ewsGetCreds();
    if (creds.ewsUrl && !_ewsMeetings.length) {
      window.syncOutlookCalendar(document.getElementById("cal-sync-btn"));
    }
  }
};

// ── Clock ──────────────────────────────────────────────────────
function tick() {
  const n = new Date();
  const clockEl = document.getElementById("clk");
  const dateEl  = document.getElementById("dt-lbl");
  const tbTime  = document.getElementById("topbar-time");
  if (clockEl) clockEl.textContent = n.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
  if (dateEl)  dateEl.textContent  = n.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
  if (tbTime)  tbTime.textContent  = n.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" });
  checkAndShowReminder(events);
}
setInterval(tick, 1000);
tick();

// ── Calendar ───────────────────────────────────────────────────
window.calNav      = (dir) => { _calNav(dir, events); };
window.addEvent    = ()    => { events = _addEvent(events); renderCalendar(events); renderSchedule(todayKey(), events); renderOvMeetings(events); document.getElementById("cal-badge").textContent = (events[todayKey()] || []).length; };

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
    if (statusEl) statusEl.textContent = "Saved locally — no SharePoint token. Connect Microsoft 365 in Weekly Report → Settings to upload.";
    return;
  }

  const fd = new FormData();
  fd.append("file",       file);
  fd.append("token",      token);
  fd.append("folderPath", path);
  const data = await api.uploadDoc(fd);

  if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-upload"></i> Upload'; }

  if (data.error) {
    if (statusEl) statusEl.textContent = "Upload failed: " + data.error;
    showToast("Upload failed: " + data.error);
    return;
  }

  DOCS[0].url = data.url || "";
  DOCS[0].s   = "Approved";
  filterDocs();

  // Reset form
  fileInput.value = "";
  const lbl = document.getElementById("dc-file-name");
  if (lbl) lbl.textContent = "Choose file (.pdf / .docx)";
  if (statusEl) statusEl.textContent = `Uploaded: ${data.name} → ${path}`;
  showToast(`Uploaded: ${data.name}`);
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
  const tok    = getMsToken();
  const badge  = document.getElementById("tm-connected-badge");
  const hint   = document.getElementById("ol-token-hint");
  const ok     = document.getElementById("ol-token-ok");
  const inp    = document.getElementById("ms-token-input");
  if (badge) badge.style.display = tok ? "inline-flex" : "none";
  if (hint)  hint.style.display  = tok ? "none"        : "inline";
  if (ok)    ok.style.display    = tok ? "inline"      : "none";
  if (inp && !inp.value && tok)  inp.value = tok;
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
    row.onclick = () => showICSMeetingDetail(m, i, row);
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
function ewsRestoreUI() {
  const { ewsUrl, username } = ewsGetCreds();
  if (!ewsUrl || !username) return;
  document.getElementById("ews-setup-form").style.display     = "none";
  document.getElementById("ews-connected-state").style.display = "block";
  document.getElementById("ews-status-badge").style.display   = "inline-flex";
  document.getElementById("ews-notconn-badge").style.display  = "none";
  const userLbl = document.getElementById("ews-user-label");
  if (userLbl) userLbl.textContent = username + "  ·  " + ewsUrl.replace("https://","").replace("/EWS/Exchange.asmx","");
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
    row.onclick = () => showEWSMeetingDetail(m, i, row);
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

window.ewsConnect = async function (btn) {
  const url      = document.getElementById("ews-url")?.value?.trim();
  const username = document.getElementById("ews-username")?.value?.trim();
  const password = document.getElementById("ews-password")?.value?.trim();
  const errEl    = document.getElementById("ews-error");

  if (!url || !username || !password) {
    if (errEl) { errEl.textContent = "Please fill in the server URL, username, and password."; errEl.style.display = "block"; }
    return;
  }
  if (errEl) errEl.style.display = "none";

  btn.disabled = true;
  btn.innerHTML = '<i class="ti ti-loader" style="font-size:12px;animation:spin .7s linear infinite"></i> Connecting...';

  const data = await api.ewsMeetings({ ewsUrl: url, username, password });

  btn.disabled = false;
  btn.innerHTML = '<i class="ti ti-plug"></i> Connect &amp; sync';

  if (data.error) {
    if (errEl) { errEl.textContent = data.error; errEl.style.display = "block"; }
    return;
  }

  // Save credentials after successful connection
  ewsSaveCreds(url, username, password);
  ewsRestoreUI();

  _ewsMeetings = data.meetings || [];
  renderEWSMeetings(_ewsMeetings);

  const syncLbl = document.getElementById("ews-last-sync-lbl");
  if (syncLbl) syncLbl.textContent = "Last synced: " + new Date().toLocaleTimeString("en-IN");

  showToast(`Connected to Exchange. Loaded ${_ewsMeetings.length} meeting(s).`);
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

  const syncLbl = document.getElementById("ews-last-sync-lbl");
  if (syncLbl) syncLbl.textContent = "Last synced: " + new Date().toLocaleTimeString("en-IN");

  showToast(`Synced ${_ewsMeetings.length} meeting(s) from Exchange.`);
};

window.ewsDisconnect = function () {
  localStorage.removeItem("ewsUrl");
  localStorage.removeItem("ewsUsername");
  localStorage.removeItem("ewsPassword");
  if (_ewsAutoSyncTimer) { clearInterval(_ewsAutoSyncTimer); _ewsAutoSyncTimer = null; }

  document.getElementById("ews-setup-form").style.display     = "block";
  document.getElementById("ews-connected-state").style.display = "none";
  document.getElementById("ews-status-badge").style.display   = "none";
  document.getElementById("ews-notconn-badge").style.display  = "inline";

  _ewsMeetings = [];
  const list = document.getElementById("teams-list");
  if (list) list.innerHTML = '<div style="font-size:12px;color:var(--hint);padding:20px 0;text-align:center"><i class="ti ti-calendar-search" style="display:block;font-size:24px;margin-bottom:8px;color:var(--border)"></i>Connect to Exchange above — meetings will load automatically.</div>';
  const countEl = document.getElementById("meetings-count");
  if (countEl) countEl.style.display = "none";
  showToast("Disconnected from Exchange.");
};

window.ewsToggleAutoSync = function (checkbox) {
  if (checkbox.checked) {
    _ewsAutoSyncTimer = setInterval(() => window.ewsSync(null), 5 * 60 * 1000);
    showToast("Auto-sync enabled — will refresh every 5 minutes.");
  } else {
    if (_ewsAutoSyncTimer) { clearInterval(_ewsAutoSyncTimer); _ewsAutoSyncTimer = null; }
    showToast("Auto-sync disabled.");
  }
};

window.ewsDiscover = async function (btn) {
  const username = document.getElementById("ews-username")?.value?.trim();
  const email    = username || "aditya.sridhar@esds.co.in";
  const resDiv   = document.getElementById("ews-discover-results");
  if (!resDiv) return;
  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader" style="font-size:12px;animation:spin .7s linear infinite"></i> Detecting...'; }

  const data = await api.ewsDiscover(email);

  if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-search"></i> Auto-detect server URL'; }

  if (data.error) { showToast(data.error); return; }

  const candidates = data.candidates || [];
  resDiv.style.display = "block";
  resDiv.innerHTML = `<div style="font-size:12px;color:var(--muted);margin-bottom:6px">Common EWS URLs for <strong>${escHtml(data.domain)}</strong> — click one to use it:</div>` +
    candidates.map((c) => `<div style="padding:5px 0;border-bottom:1px solid var(--border);font-size:12px;font-family:var(--mono);display:flex;align-items:center;justify-content:space-between;gap:8px"><span style="color:var(--text);word-break:break-all">${escHtml(c)}</span><button class="sm" style="flex-shrink:0" onclick="document.getElementById('ews-url').value='${escHtml(c)}';document.getElementById('ews-discover-results').style.display='none'">Use</button></div>`).join("");
};

// ── Teams MOM ──────────────────────────────────────────────────
const TEAMS_MEETINGS = [
  { t: "EEL architecture review", d: "03 Jun 2026", dur: "45 min", att: "Igor, Tushar, Aditya S, Vijay", ctx: "EEL OS layer BRD confirmation, VXLAN Phase 2, Sub manager POC deadline 15 Jun" },
  { t: "GPOS sub manager demo",   d: "02 Jun 2026", dur: "60 min", att: "Sahil, Aditya S, Vijay, Igor",  ctx: "Full demo walkthrough, candlepin integration, SSL cert fix, Phase 3 scope discussion" },
  { t: "Infrastructure blueprint",d: "01 Jun 2026", dur: "30 min", att: "Tushar, Aditya S, Igor",        ctx: "AHCP blueprint review, network topology, rack layout approval" },
  { t: "Weekly stand-up",         d: "09 Jun 2026", dur: "15 min", att: "Full team",                     ctx: "Status updates, SIT blocker raised, UAT prep timeline discussed" },
];
window.selectMeeting = function (el, i) {
  document.querySelectorAll(".meet-row").forEach((r) => r.classList.remove("sel"));
  el.classList.add("sel");
  const m   = TEAMS_MEETINGS[i];
  const det = document.getElementById("meeting-detail");
  if (!det) return;
  det.innerHTML = `<div style="font-size:14px;font-weight:500;margin-bottom:5px">${m.t}</div><div style="font-size:12px;color:var(--muted);margin-bottom:7px">${m.d} · ${m.dur} · ${m.att}</div><div style="font-size:13px;color:var(--muted);margin-bottom:11px">${m.ctx}</div><button class="primary" onclick="generateTeamsMOM(${i})"><i class="ti ti-bolt"></i> Generate MOM</button>`;
};
window.generateTeamsMOM = async function (i) {
  const m   = TEAMS_MEETINGS[i];
  const out = document.getElementById("mom-tm-output");
  const txt = document.getElementById("mom-tm-text");
  if (out) out.style.display = "block";
  if (txt) txt.innerHTML = `<span class="ldot" style="background:var(--blue)"></span>Generating MOM...`;
  const data = await api.generateMom({ title: m.t, date: m.d, attendees: m.att, objective: m.ctx });
  if (txt) typeIn(txt, data.text || data.error || "Error generating MOM");
};
window.refreshMeetings = function (btn) {
  if (!btn) return;
  btn.innerHTML = '<i class="ti ti-refresh" style="font-size:13px;animation:spin .7s linear infinite"></i>';
  setTimeout(() => btn.innerHTML = '<i class="ti ti-refresh" style="font-size:13px"></i>', 1200);
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
    row.onclick   = () => selectMeeting(row, -1, m);
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
    <div id="ol-body-content" style="font-size:13px;color:var(--text);white-space:pre-wrap;line-height:1.7;min-height:60px;margin-bottom:12px">
      <span class="ldot" style="background:var(--blue)"></span> Loading full email...
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
    const bodyText = bodyData.body || "(no content)";
    const bodyEl   = document.getElementById("ol-body-content");
    if (bodyEl) bodyEl.textContent = bodyText;
    _currentEmail.bodyFull    = bodyText;
    _currentEmail.bodyPreview = bodyText.slice(0, 500);
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

window.sendComposedEmail = async function (btn) {
  const to      = document.getElementById("ol-compose-to")?.value?.trim();
  const subject = document.getElementById("ol-compose-subject")?.value?.trim();
  const body    = document.getElementById("ol-compose-body")?.value?.trim();
  if (!to)      { showToast("Enter a recipient email."); return; }
  if (!subject) { showToast("Enter a subject."); return; }

  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader" style="font-size:12px;animation:spin .7s linear infinite"></i> Sending...'; }

  const creds = ewsGetCreds();
  const data  = creds.ewsUrl
    ? await api.ewsSendEmail({ ...creds, to, subject, body })
    : await api.outlookSend({ to, subject, body });

  if (btn) { btn.disabled = false; btn.innerHTML = '<i class="ti ti-send"></i> Send'; }

  if (data.ok) {
    showToast("Email sent!");
    const area = document.getElementById("ol-compose-area");
    if (area) area.style.display = "none";
  } else {
    showToast("Send failed: " + (data.error || "Unknown error"));
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
function previewEmail(email, row) {
  document.querySelectorAll("#ol-list .meet-row").forEach((r) => r.classList.remove("sel"));
  row.classList.add("sel");
  _currentEmail = email;
  const det = document.getElementById("ol-detail");
  if (!det) return;
  const from = email.from?.emailAddress?.name || email.from?.emailAddress?.address || "Unknown";
  det.innerHTML = `<div style="font-size:14px;font-weight:500;margin-bottom:4px">${escHtml(email.subject || "(no subject)")}</div><div style="font-size:12px;color:var(--muted);margin-bottom:8px">From: ${escHtml(from)} · ${new Date(email.receivedDateTime).toLocaleString("en-IN")}</div><div style="font-size:13px;color:var(--muted);margin-bottom:10px">${escHtml(email.bodyPreview || "")}</div><button class="primary sm" onclick="draftEmailReply()"><i class="ti ti-robot"></i> AI Draft Reply</button>`;
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
  const to = _currentEmail.from?.emailAddress?.address  // M365
           || _currentEmail.from?.address;              // EWS
  if (!to) { showToast("Cannot determine reply-to address."); return; }
  const data = await api.outlookSend({ to, subject: "Re: " + (_currentEmail.subject || ""), body: txt });
  showToast(data.ok ? "Email sent!" : "Send failed: " + data.error);
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
  const statusEl = document.getElementById("cal-sync-status");
  if (statusEl) statusEl.textContent = `Last synced: ${new Date().toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" })} · ${_ewsMeetings.length} event(s)`;
  showToast(`Calendar synced — ${_ewsMeetings.length} event(s) from Exchange.`);
};

// ── Weekly report ──────────────────────────────────────────────
window.wrTab = function (el, tab) {
  document.querySelectorAll(".seg button").forEach((b) => b.classList.remove("on"));
  el.classList.add("on");
  ["gen", "sp", "cfg"].forEach((t) => { const el = document.getElementById("wr-" + t); if (el) el.style.display = "none"; });
  const active = document.getElementById("wr-" + tab);
  if (active) active.style.display = "block";
};
window.generateReport = async function () {
  const t = document.getElementById("rp-output");
  if (!t) return;
  t.innerHTML = `<span class="ldot" style="background:var(--blue)"></span>Scanning AI history and generating report...`;
  const data = await api.generateReport({
    name:    document.getElementById("rp-name")?.value,
    dept:    document.getElementById("rp-dept")?.value,
    week:    document.getElementById("rp-week")?.value,
    manager: document.getElementById("rp-mgr")?.value,
    source:  document.getElementById("rp-source")?.value,
  });
  typeIn(t, data.text || data.error || "Error generating report");
};
window.toggleOAuth = function () {
  const g = document.getElementById("oauth-steps");
  if (g) g.style.display = g.style.display === "none" ? "block" : "none";
};
window.loginMicrosoft = function () {
  window.location.href = "/api/auth/microsoft";
};
window.testSharePoint = async function () {
  const s   = document.getElementById("sp-status");
  const tok = document.getElementById("sp-token")?.value || localStorage.getItem("spToken") || "";
  if (!s) return;
  s.style.display = "block";
  if (!tok) { typeIn(s, "No access token. Click OAuth setup → Login with Microsoft 365."); return; }
  s.innerHTML = `<span class="ldot" style="background:var(--blue)"></span>Testing connection...`;
  const data = await api.spTest(tok);
  typeIn(s, data.ok ? `✓ Connected\nUser : ${data.user} (${data.email})\nDrive: ${data.driveType}\nStatus: Ready` : `✗ Failed: ${data.error}`);
};
window.exportToSharePoint = async function () {
  const s       = document.getElementById("sp-status");
  const content = document.getElementById("rp-output")?.innerText || "";
  if (!s) return;
  if (!content || content.includes("Click")) { showToast("Generate a report first, then export."); return; }
  s.style.display = "block";
  s.innerHTML = `<span class="ldot" style="background:var(--blue)"></span>Building .docx and uploading...`;
  const data = await api.spExport({
    token:         document.getElementById("sp-token")?.value || localStorage.getItem("spToken") || "",
    spUrl:         document.getElementById("sp-url")?.value,
    filename:      document.getElementById("sp-filename")?.value,
    source:        document.getElementById("sp-source")?.value,
    reportContent: content,
    reportMeta: {
      name:    document.getElementById("rp-name")?.value,
      dept:    document.getElementById("rp-dept")?.value,
      week:    document.getElementById("rp-week")?.value,
      manager: document.getElementById("rp-mgr")?.value,
    },
  });
  typeIn(s, data.ok
    ? `✓ Exported!\nFile   : ${data.path || ""}\nURL    : ${data.webUrl || "—"}\nTime   : ${new Date().toLocaleString("en-IN")}`
    : `✗ Failed: ${data.error}`
  );
};
window.saveKeys = function () {
  const ak = document.getElementById("cfg-ak")?.value;
  const ok = document.getElementById("cfg-ok")?.value;
  if (ak) localStorage.setItem("anthropicKey", ak);
  if (ok) localStorage.setItem("openaiKey", ok);
  showToast("API keys saved.");
};

// ── MOM generator ──────────────────────────────────────────────
window.generateMOM = async function () {
  const res = document.getElementById("mom-result");
  const o   = document.getElementById("mom-output");
  if (res) res.style.display = "block";
  if (o)   o.innerHTML = `<span class="ldot" style="background:var(--blue)"></span>Generating MOM...`;
  const data = await api.generateMom({
    title:      document.getElementById("mm-title")?.value,
    date:       document.getElementById("mm-date")?.value,
    attendees:  document.getElementById("mm-att")?.value,
    facilitator:document.getElementById("mm-fac")?.value,
    objective:  document.getElementById("mm-obj")?.value,
    transcript: document.getElementById("mm-transcript")?.value,
  });
  if (o) typeIn(o, data.text || data.error);
};
window.clearMOM = function () {
  ["mm-title","mm-date","mm-att","mm-obj","mm-transcript"].forEach((id) => { const e = document.getElementById(id); if (e) e.value = ""; });
  const fac = document.getElementById("mm-fac"); if (fac) fac.value = "Aditya S";
  const res = document.getElementById("mom-result"); if (res) res.style.display = "none";
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
  const done     = document.getElementById("su-done")?.value    || "";
  const today    = document.getElementById("su-today")?.value   || "";
  const blockers = document.getElementById("su-blockers")?.value || "";
  const format   = document.getElementById("su-format")?.value  || "";
  const data = await api.generateStandup({ done, today, blockers, format, name: "Aditya S" });
  typeIn(t, data.text || data.error);
  if (data.text) {
    _standupCtx = { done, today, blockers, standupText: data.text };
    const qaSection = document.getElementById("su-qa-section");
    if (qaSection) {
      document.getElementById("su-qa-history").innerHTML = "";
      qaSection.style.display = "block";
      _suRenderSuggestions();
    }
  }
};
window.loadYesterday = function () {
  const done  = document.getElementById("su-done");
  const today = document.getElementById("su-today");
  if (done)  done.value  = "Completed EEL Bug Portal customer & admin console. Applied ESDS doc policy to GPOS BRD v1.1/v1.2.";
  if (today) today.value = "Working on SharePoint export integration for weekly reports. Reviewing GPOS BRD v1.2 changes.";
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

  const data = await api.standupQA({ ..._standupCtx, question });

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

window.generateLeaveEmail = async function () {
  const t = document.getElementById("lv-output");
  if (!t) return;
  t.style.display = "block";
  t.innerHTML = `<span class="ldot" style="background:var(--blue)"></span>Generating leave email...`;
  const leaveType = document.getElementById("lv-type")?.value || "Planned leave";
  const fromDate  = document.getElementById("lv-from")?.value || "";
  const toDate    = document.getElementById("lv-to")?.value || "";
  const reason    = document.getElementById("lv-reason")?.value?.trim() || "";
  const toEmail   = document.getElementById("lv-to-email")?.value?.trim() || "";
  const fmt = (d) => d ? new Date(d).toLocaleDateString("en-IN", { day: "2-digit", month: "long", year: "numeric" }) : d;
  const dateRange = fromDate === toDate ? fmt(fromDate) : `${fmt(fromDate)} to ${fmt(toDate)}`;
  const prompt = `Write a formal, professional ${leaveType.toLowerCase()} request email from Aditya S (System Analyst / Solution Architect, ESDS Software Solution Pvt. Ltd.)` +
    ` for dates: ${dateRange}.` +
    (reason ? ` Reason: ${reason}.` : "") +
    ` Include: greeting, leave dates, brief reason, assurance of handover, request for approval, and thank you. Keep it brief and professional. Do not include subject line.`;
  const data = await api.chat(prompt, "You write formal, courteous workplace emails.");
  typeIn(t, data.text || data.error);
};

window.sendLeaveEmail = async function (btn) {
  const emailText = document.getElementById("lv-output")?.innerText?.trim();
  if (!emailText || emailText.includes("Generating")) {
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
window.sendChatMessage = async function () {
  const inp = document.getElementById("chat-input");
  const msg = inp?.value?.trim();
  if (!msg) return;
  if (inp) inp.value = "";
  const msgs = document.getElementById("chat-messages");
  if (!msgs) return;
  msgs.innerHTML += `<div class="chat-msg user"><div class="msg-lbl">You</div>${escHtml(msg)}</div>`;
  const thinking = document.createElement("div");
  thinking.className = "chat-msg ai";
  thinking.innerHTML = `<div class="msg-lbl">Assistant</div><span class="ldot" style="background:var(--muted)"></span>Thinking...`;
  msgs.appendChild(thinking);
  msgs.scrollTop = msgs.scrollHeight;
  const data = await api.chat(msg);
  thinking.innerHTML = `<div class="msg-lbl">Assistant</div>${escHtml(data.text || data.error || "Error").replace(/\n/g,"<br>")}`;
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
  if (msgs) msgs.innerHTML = '<div class="chat-msg ai"><div class="msg-lbl">Assistant</div>Chat cleared. How can I help?</div>';
};

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
  // Calendar
  renderCalendar(events);
  renderSchedule(todayKey(), events);
  renderOvMeetings(events);
  const badgeEl = document.getElementById("cal-badge");
  if (badgeEl) badgeEl.textContent = (events[todayKey()] || []).length;

  // Charts (defer so Canvas is ready)
  setTimeout(() => {
    healthChart = initHealthChart();
    initDocChart();
  }, 300);

  // Documents
  renderDocs(window._DOCS || []);  // renderDocs imported at top of documents.js
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

  // Outlook EWS status
  const ewsCreds = ewsGetCreds();
  const olHint   = document.getElementById("ol-ews-hint");
  const olOk     = document.getElementById("ol-ews-ok");
  if (ewsCreds.ewsUrl && olHint && olOk) { olHint.style.display = "none"; olOk.style.display = "inline"; }

  // Restore EWS connection and auto-sync if previously connected
  const { ewsUrl, username, password } = ewsCreds;
  if (ewsUrl && username && password) {
    ewsRestoreUI();
    // Auto-load meetings in background on startup
    api.ewsMeetings({ ewsUrl, username, password }).then((data) => {
      if (data.error) return;
      _ewsMeetings = data.meetings || [];
      renderEWSMeetings(_ewsMeetings);
      const syncLbl = document.getElementById("ews-last-sync-lbl");
      if (syncLbl) syncLbl.textContent = "Last synced: " + new Date().toLocaleTimeString("en-IN");
    });
  }
});
