// frontend/src/pages/overview.js
import { todayKey } from "../utils/calendar.js";

export function renderOvMeetings(events) {
  const evts = (events[todayKey()] || []).filter(
    (e) => e.type === "meet" || e.type === "review"
  );
  const countEl = document.getElementById("today-count");
  const listEl  = document.getElementById("ov-meetings");
  if (countEl) countEl.textContent = `${evts.length} meeting${evts.length !== 1 ? "s" : ""}`;
  if (!listEl) return;

  if (!evts.length) {
    listEl.innerHTML = '<div style="font-size:13px;color:var(--muted);padding:8px 0">No meetings scheduled today</div>';
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

export function initHealthChart() {
  const canvas = document.getElementById("health-chart");
  if (!canvas || typeof Chart === "undefined") return null;
  const dk = matchMedia("(prefers-color-scheme:dark)").matches;
  return new Chart(canvas, {
    type: "bar",
    data: {
      labels: ["EEL OS", "GPOS", "AIOps", "Bug Portal"],
      datasets: [{
        data: [82, 71, 45, 60],
        backgroundColor: ["#059669", "#1d6ae5", "#d97706", "#1d6ae5"],
        borderRadius: 4,
        borderSkipped: false,
      }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 10 }, color: dk ? "#8b9ab0" : "#6b7280" } },
        y: { max: 100, grid: { color: dk ? "rgba(255,255,255,.07)" : "rgba(0,0,0,.05)" }, ticks: { font: { size: 10 }, color: dk ? "#8b9ab0" : "#6b7280", callback: (v) => v + "%" } },
      },
    },
  });
}

export function initDocChart() {
  const canvas = document.getElementById("doc-chart");
  if (!canvas || typeof Chart === "undefined") return null;
  return new Chart(canvas, {
    type: "doughnut",
    data: {
      labels: ["Approved", "Pending", "Overdue"],
      datasets: [{ data: [13, 4, 1], backgroundColor: ["#059669", "#d97706", "#dc2626"], borderWidth: 0 }],
    },
    options: { responsive: true, maintainAspectRatio: false, cutout: "70%", plugins: { legend: { display: false } } },
  });
}

export function filterHealthChart(chart, value) {
  const maps = {
    all:  { l: ["EEL OS", "GPOS", "AIOps", "Bug Portal"], d: [82, 71, 45, 60], c: ["#059669","#1d6ae5","#d97706","#1d6ae5"] },
    eel:  { l: ["EEL OS", "Bug Portal"], d: [82, 60], c: ["#059669","#1d6ae5"] },
    gpos: { l: ["GPOS"], d: [71], c: ["#1d6ae5"] },
  };
  const m = maps[value] || maps.all;
  chart.data.labels = m.l;
  chart.data.datasets[0].data = m.d;
  chart.data.datasets[0].backgroundColor = m.c;
  chart.update();
  ["ph-gpos","ph-aiops","ph-bug"].forEach((id, i) => {
    const el = document.getElementById(id);
    if (el) el.style.display = value === "eel" && i <= 1 ? "none" : value === "gpos" && i !== 0 ? "none" : "";
  });
}
