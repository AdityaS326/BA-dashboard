// frontend/src/utils/calendar.js
// Calendar state management and helper functions.

export const MONTH_NAMES = [
  "January","February","March","April","May","June",
  "July","August","September","October","November","December",
];
export const DAY_NAMES = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];

export function dayKey(y, m, d) {
  return `${y}-${String(m + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}

export function todayKey() {
  const n = new Date();
  return dayKey(n.getFullYear(), n.getMonth(), n.getDate());
}

// Returns an empty events object — all real events come from EWS via syncOutlookCalendar()
export function seedEvents() {
  return {};
}

export function checkUpcomingReminder(events) {
  const tk = todayKey();
  const evts = events[tk] || [];
  const now  = new Date();
  const nowMin = now.getHours() * 60 + now.getMinutes();
  let next = null, diff = Infinity;
  evts.forEach((ev) => {
    const d = ev.h * 60 + ev.m - nowMin;
    if (d > 0 && d < diff) { diff = d; next = ev; }
  });
  return next && diff <= 30 ? { event: next, minutesAway: diff } : null;
}
