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

// Default events seeded for demo
export function seedEvents() {
  const tk = todayKey();
  const events = {};

  events[tk] = [
    { title: "Daily stand-up",         time: "09:00 AM", type: "meet",     att: "Full team",                    dur: "15 min", h: 9,  m: 0  },
    { title: "EEL Architecture Review", time: "10:30 AM", type: "meet",     att: "Igor, Tushar, Aditya, Vijay", dur: "45 min", h: 10, m: 30 },
    { title: "GPOS BRD review",         time: "02:00 PM", type: "review",   att: "Igor, Aditya",                dur: "60 min", h: 14, m: 0  },
    { title: "SIT status check",        time: "04:00 PM", type: "meet",     att: "Tushar, Aditya",              dur: "30 min", h: 16, m: 0  },
  ];

  // Tomorrow
  const d2 = new Date(); d2.setDate(d2.getDate() + 1);
  const k2 = dayKey(d2.getFullYear(), d2.getMonth(), d2.getDate());
  events[k2] = [
    { title: "Infrastructure review", time: "11:00 AM", type: "meet",     att: "Tushar, Igor", dur: "30 min", h: 11, m: 0 },
    { title: "FRD deadline",          time: "06:00 PM", type: "deadline", att: "Aditya S",     dur: "—",      h: 18, m: 0 },
  ];

  // 3 days from now
  const d3 = new Date(); d3.setDate(d3.getDate() + 3);
  const k3 = dayKey(d3.getFullYear(), d3.getMonth(), d3.getDate());
  events[k3] = [
    { title: "Weekly review", time: "10:00 AM", type: "meet", att: "Full team", dur: "45 min", h: 10, m: 0 },
  ];

  return events;
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
