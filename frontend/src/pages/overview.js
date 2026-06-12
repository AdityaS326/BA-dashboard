// frontend/src/pages/overview.js
import { todayKey, MONTH_NAMES } from "../utils/calendar.js";
import { DOCS } from "./documents.js";

// ── Today's meetings in overview ──────────────────────────────
export function renderOvMeetings(events) {
  const evts    = (events[todayKey()] || []).filter((e) => e.type === "meet" || e.type === "review");
  const countEl = document.getElementById("today-count");
  const listEl  = document.getElementById("ov-meetings");
  if (countEl) countEl.textContent = evts.length ? `${evts.length} meeting${evts.length !== 1 ? "s" : ""}` : "—";
  if (!listEl) return;

  if (!evts.length) {
    listEl.innerHTML = '<div style="font-size:13px;color:var(--muted);padding:16px 0;text-align:center"><i class="ti ti-plug" style="display:block;font-size:22px;margin-bottom:6px;color:var(--border)"></i>Connect Exchange in Teams panel to see today\'s meetings</div>';
    return;
  }

  listEl.innerHTML = "";
  evts.forEach((ev) => {
    const row = document.createElement("div");
    row.style.cssText = "display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid var(--border);cursor:pointer";
    row.innerHTML = `
      <div style="width:38px;height:38px;border-radius:var(--r-sm);background:var(--blue-bg);display:flex;align-items:center;justify-content:center;flex-shrink:0">
        <i class="ti ti-calendar-event" style="font-size:17px;color:var(--blue)"></i>
      </div>
      <div style="flex:1;min-width:0">
        <div style="font-size:13px;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${ev.title}</div>
        <div style="font-size:11px;color:var(--muted);margin-top:1px">${ev.time} · ${ev.dur}</div>
      </div>`;
    row.onclick = () => document.querySelector('[data-panel="cal"]')?.click();
    listEl.appendChild(row);
  });
  if (listEl.lastChild) listEl.lastChild.style.borderBottom = "none";
}

// ── Upcoming meetings (next 5) ─────────────────────────────────
export function renderOvUpcoming(ewsMeetings) {
  const listEl  = document.getElementById("ov-upcoming");
  const countEl = document.getElementById("upcoming-count");
  if (!listEl) return;

  const now      = new Date();
  const upcoming = (ewsMeetings || [])
    .filter((m) => new Date(m.start) > now)
    .sort((a, b) => new Date(a.start) - new Date(b.start))
    .slice(0, 5);

  if (countEl) countEl.textContent = upcoming.length || "—";

  if (!upcoming.length) {
    listEl.innerHTML = '<div style="font-size:13px;color:var(--muted);padding:16px 0;text-align:center"><i class="ti ti-plug" style="display:block;font-size:22px;margin-bottom:6px;color:var(--border)"></i>No upcoming meetings — sync Exchange</div>';
    return;
  }

  listEl.innerHTML = "";
  upcoming.forEach((m) => {
    const start   = new Date(m.start);
    const dateStr = isNaN(start) ? "" : start.toLocaleDateString("en-IN", { weekday: "short", day: "2-digit", month: "short" });
    const timeStr = isNaN(start) ? "" : start.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" });
    const row = document.createElement("div");
    row.style.cssText = "display:flex;align-items:center;gap:10px;padding:7px 0;border-bottom:1px solid var(--border)";
    row.innerHTML = `
      <div style="width:36px;height:36px;border-radius:var(--r-sm);background:var(--surface2);display:flex;flex-direction:column;align-items:center;justify-content:center;flex-shrink:0;font-size:9px;color:var(--muted);font-weight:500;text-transform:uppercase;line-height:1.3">
        <span>${isNaN(start) ? "?" : start.getDate()}</span>
        <span>${isNaN(start) ? "" : MONTH_NAMES[start.getMonth()].slice(0, 3)}</span>
      </div>
      <div style="flex:1;min-width:0">
        <div style="font-size:12px;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${m.subject || "Untitled"}</div>
        <div style="font-size:11px;color:var(--muted);margin-top:1px">${dateStr} · ${timeStr} · ${m.dur || ""}</div>
      </div>
      ${m.isOnline ? '<span class="badge b-blue" style="flex-shrink:0;font-size:10px">Teams</span>' : ""}`;
    row.onclick = () => document.querySelector('[data-panel="tm"]')?.click();
    listEl.appendChild(row);
  });
  if (listEl.lastChild) listEl.lastChild.style.borderBottom = "none";
}

// ── Overview stat cards (all from live EWS + DOCS) ────────────
export function updateOverviewStats(ewsMeetings, inboxCount) {
  const now       = new Date();
  const todayStr  = now.toDateString();

  // Meetings today
  const todayCount = (ewsMeetings || []).filter((m) => new Date(m.start).toDateString() === todayStr).length;
  const todayEl    = document.getElementById("stat-today-meetings");
  const todaySubEl = document.getElementById("stat-today-sub");
  if (todayEl) todayEl.textContent = todayCount;
  if (todaySubEl) todaySubEl.textContent = todayCount === 1 ? "1 meeting today" : `${todayCount} meetings today`;

  // Meetings this week
  const weekStart  = new Date(now); weekStart.setDate(now.getDate() - now.getDay()); weekStart.setHours(0,0,0,0);
  const weekEnd    = new Date(weekStart); weekEnd.setDate(weekStart.getDate() + 7);
  const weekCount  = (ewsMeetings || []).filter((m) => { const d = new Date(m.start); return d >= weekStart && d < weekEnd; }).length;
  const weekEl     = document.getElementById("stat-week-meetings");
  const weekSubEl  = document.getElementById("stat-week-sub");
  if (weekEl) weekEl.textContent = weekCount;
  if (weekSubEl) weekSubEl.textContent = `meetings this week`;

  // Documents
  const docCount   = DOCS.length;
  const docEl      = document.getElementById("stat-docs");
  const docSubEl   = document.getElementById("stat-docs-sub");
  if (docEl) docEl.textContent = docCount;
  if (docSubEl) docSubEl.textContent = docCount ? `${DOCS.filter((d) => d.s === "Approved").length} approved` : "Add in Documents tab";

  // Unread emails
  const unreadEl    = document.getElementById("stat-unread");
  const unreadSubEl = document.getElementById("stat-unread-sub");
  if (unreadEl) unreadEl.textContent = inboxCount != null ? inboxCount : "—";
  if (unreadSubEl) unreadSubEl.textContent = inboxCount != null ? (inboxCount === 1 ? "1 unread" : `${inboxCount} unread`) : "Connect Exchange";

  // Meeting activity bars
  renderMeetingActivity(ewsMeetings);
}

// ── Meeting activity chart ─────────────────────────────────────
function renderMeetingActivity(ewsMeetings) {
  const wrap = document.getElementById("meeting-activity");
  if (!wrap) return;
  if (!ewsMeetings || !ewsMeetings.length) return;

  const days = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
  const counts = [0,0,0,0,0,0,0];
  const now = new Date();
  const weekStart = new Date(now); weekStart.setDate(now.getDate() - now.getDay()); weekStart.setHours(0,0,0,0);
  const weekEnd = new Date(weekStart); weekEnd.setDate(weekStart.getDate() + 7);

  ewsMeetings.forEach((m) => {
    const d = new Date(m.start);
    if (d >= weekStart && d < weekEnd) counts[d.getDay()]++;
  });

  const max = Math.max(...counts, 1);
  wrap.innerHTML = `<div style="display:flex;align-items:flex-end;gap:5px;height:64px;padding:4px 0">` +
    days.map((day, i) => {
      const h = Math.round((counts[i] / max) * 52);
      const isToday = i === now.getDay();
      return `<div style="flex:1;display:flex;flex-direction:column;align-items:center;gap:3px">
        <div style="font-size:11px;color:var(--muted);font-weight:${counts[i] ? "500" : "400"}">${counts[i] || ""}</div>
        <div style="width:100%;height:${h || 4}px;background:${isToday ? "var(--blue)" : "var(--surface2)"};border-radius:3px;border:1px solid ${isToday ? "var(--blue)" : "var(--border)"}"></div>
        <div style="font-size:10px;color:${isToday ? "var(--blue)" : "var(--muted)"};font-weight:${isToday ? "600" : "400"}">${day}</div>
      </div>`;
    }).join("") + `</div>`;
}

// ── Document donut chart ───────────────────────────────────────
let _docChart = null;

export function initDocChart() {
  const canvas = document.getElementById("doc-chart");
  if (!canvas || typeof Chart === "undefined") return null;
  _docChart = new Chart(canvas, {
    type: "doughnut",
    data: {
      labels: ["Approved", "Pending/Draft", "Overdue"],
      datasets: [{ data: [0, 0, 0], backgroundColor: ["#059669", "#d97706", "#dc2626"], borderWidth: 0 }],
    },
    options: { responsive: true, maintainAspectRatio: false, cutout: "70%", plugins: { legend: { display: false } } },
  });
  return _docChart;
}

export function updateDocChart() {
  const approved = DOCS.filter((d) => d.s === "Approved" || d.s === "Delivered").length;
  const pending  = DOCS.filter((d) => d.s === "Pending"  || d.s === "Draft").length;
  const overdue  = DOCS.filter((d) => d.s === "Overdue").length;

  const aEl = document.getElementById("doc-count-approved");
  const pEl = document.getElementById("doc-count-pending");
  const oEl = document.getElementById("doc-count-overdue");
  if (aEl) aEl.textContent = approved;
  if (pEl) pEl.textContent = pending;
  if (oEl) oEl.textContent = overdue;

  if (_docChart) {
    _docChart.data.datasets[0].data = [approved, pending, overdue];
    _docChart.update();
  }
}

// ── Health chart stub (kept for compatibility) ─────────────────
export function initHealthChart() { return null; }
export function filterHealthChart() {}
