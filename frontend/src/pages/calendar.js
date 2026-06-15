// frontend/src/pages/calendar.js
import { MONTH_NAMES, DAY_NAMES, dayKey, todayKey, checkUpcomingReminder } from "../utils/calendar.js";
import { showToast } from "../utils/ui.js";

let calYear  = new Date().getFullYear();
let calMonth = new Date().getMonth();
let selectedDay = todayKey();

export function renderCalendar(events) {
  const monthLbl = document.getElementById("cal-month-lbl");
  if (monthLbl) monthLbl.textContent = `${MONTH_NAMES[calMonth]} ${calYear}`;

  // Headers
  const hdr = document.getElementById("cal-header");
  if (hdr) {
    hdr.innerHTML = "";
    DAY_NAMES.forEach((d) => {
      const el = document.createElement("div");
      el.className = "cal-hdr";
      el.textContent = d;
      hdr.appendChild(el);
    });
  }

  // Cells
  const cells = document.getElementById("cal-cells");
  if (!cells) return;
  cells.innerHTML = "";
  const firstDay    = new Date(calYear, calMonth, 1).getDay();
  const daysInMonth = new Date(calYear, calMonth + 1, 0).getDate();
  const prevTotal   = new Date(calYear, calMonth, 0).getDate();
  const tk          = todayKey();

  for (let i = 0; i < firstDay; i++) {
    const el = document.createElement("div");
    el.className = "cal-cell other";
    el.textContent = prevTotal - firstDay + 1 + i;
    cells.appendChild(el);
  }

  for (let d = 1; d <= daysInMonth; d++) {
    const k  = dayKey(calYear, calMonth, d);
    const el = document.createElement("div");
    let cls  = "cal-cell";
    if (k === tk)          cls += " today";
    if (events[k])         cls += " has-ev";
    if (k === selectedDay) cls += " sel";
    el.className = cls;
    el.textContent = d;
    el.onclick = () => {
      selectedDay = k;
      document.querySelectorAll(".cal-cell").forEach((c) => c.classList.remove("sel"));
      el.classList.add("sel");
      renderSchedule(k, events);
      const evDate = document.getElementById("ev-date");
      if (evDate) evDate.value = k;
    };
    cells.appendChild(el);
  }

  let nd = 1;
  while (cells.children.length % 7 !== 0) {
    const el = document.createElement("div");
    el.className = "cal-cell other";
    el.textContent = nd++;
    cells.appendChild(el);
  }
}

export function renderSchedule(k, events) {
  const evts    = events[k] || [];
  const dateLbl = document.getElementById("sch-date-lbl");
  const countEl = document.getElementById("sch-count");
  const el      = document.getElementById("day-schedule");

  const parts = k.split("-");
  if (dateLbl) dateLbl.textContent = `${parseInt(parts[2])} ${MONTH_NAMES[parseInt(parts[1]) - 1]}`;
  if (countEl) countEl.textContent = `${evts.length} event${evts.length !== 1 ? "s" : ""}`;
  if (!el) return;

  if (!evts.length) {
    el.innerHTML = `<div class="sch-empty"><i class="ti ti-calendar-off" style="font-size:22px;display:block;margin-bottom:7px;color:var(--hint)"></i>No events scheduled</div>`;
    return;
  }

  el.innerHTML = "";
  evts.forEach((ev) => {
    const slot = document.createElement("div");
    slot.className = "sch-slot";
    slot.innerHTML = `<div class="sch-time">${ev.time}</div><div class="sch-ev ${ev.type || "meet"}"><div class="sch-title">${ev.title}</div><div class="sch-meta">${ev.dur}${ev.att ? " · " + ev.att : ""}</div></div>`;
    el.appendChild(slot);
  });
}

export function calNav(dir, events) {
  calMonth += dir;
  if (calMonth > 11) { calMonth = 0; calYear++; }
  if (calMonth < 0)  { calMonth = 11; calYear--; }
  renderCalendar(events);
  renderSchedule(selectedDay, events);
}

export function addEvent(events) {
  const title = document.getElementById("ev-title")?.value.trim();
  const time  = document.getElementById("ev-time")?.value.trim();
  const dur   = document.getElementById("ev-dur")?.value;
  const att   = document.getElementById("ev-att")?.value.trim();
  const type  = document.getElementById("ev-type")?.value;

  if (!title || !time) { showToast("Please enter a title and time."); return events; }

  const m = time.match(/(\d+):(\d+)\s*(AM|PM)?/i);
  let h  = m ? parseInt(m[1]) : 9;
  let mi = m ? parseInt(m[2]) : 0;
  if (m && m[3] && m[3].toUpperCase() === "PM" && h !== 12) h += 12;
  if (m && m[3] && m[3].toUpperCase() === "AM" && h === 12) h = 0;

  const updated = { ...events };
  if (!updated[selectedDay]) updated[selectedDay] = [];
  updated[selectedDay] = [...updated[selectedDay], { title, time, type, att, dur, h, m: mi }]
    .sort((a, b) => a.h * 60 + a.m - (b.h * 60 + b.m));

  if (document.getElementById("ev-title"))    document.getElementById("ev-title").value    = "";
  if (document.getElementById("ev-att"))      document.getElementById("ev-att").value      = "";
  if (document.getElementById("ev-desc"))     document.getElementById("ev-desc").value     = "";
  if (document.getElementById("ev-location")) document.getElementById("ev-location").value = "";

  showToast("Event added to calendar.");
  return updated;
}

export function checkAndShowReminder(events) {
  const reminder = checkUpcomingReminder(events);
  const pill     = document.getElementById("reminder-pill");
  const pillTxt  = document.getElementById("reminder-txt");

  ["ov-reminder", "cal-reminder"].forEach((id) => {
    const el = document.getElementById(id);
    if (!el) return;
    if (reminder) {
      const urgent = reminder.minutesAway <= 5;
      const cls    = urgent ? "urgent" : "warn";
      const msg    = urgent
        ? `Starting now: ${reminder.event.title}`
        : `In ${reminder.minutesAway} min: ${reminder.event.title} at ${reminder.event.time}`;
      el.innerHTML = `<div class="reminder-banner ${cls}" role="alert"><i class="ti ti-bell"></i><div style="flex:1">${msg}</div><button class="sm" style="margin-left:auto;flex-shrink:0" onclick="document.querySelector('[data-panel=cal]')?.click()">View</button></div>`;
    } else {
      el.innerHTML = "";
    }
  });

  if (pill && pillTxt) {
    if (reminder) {
      const urgent = reminder.minutesAway <= 5;
      pill.style.display = "inline-flex";
      pill.className = `tb-pill ${urgent ? "urgent" : "warn"}`;
      pillTxt.textContent = `${reminder.event.title} in ${reminder.minutesAway} min`;
    } else {
      pill.style.display = "none";
    }
  }
}
