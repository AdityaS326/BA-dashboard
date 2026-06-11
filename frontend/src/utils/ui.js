// frontend/src/utils/ui.js
// Shared UI utility functions used by all page modules.

/** Animate text character-by-character into an element */
export function typeIn(el, text, speed = 7) {
  el.textContent = "";
  let i = 0;
  const iv = setInterval(() => {
    el.textContent += text[i++];
    el.scrollTop = el.scrollHeight;
    if (i >= text.length) clearInterval(iv);
  }, speed);
}

/** Copy element's inner text to clipboard and show a brief flash */
export function copyText(id) {
  const el = document.getElementById(id);
  const txt = el.innerText || el.textContent;
  navigator.clipboard.writeText(txt).then(() => showToast("Copied to clipboard.")).catch(() => {
    const ta = document.createElement("textarea");
    ta.value = txt;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    document.body.removeChild(ta);
    showToast("Copied.");
  });
}

/** Show a non-blocking toast message */
export function showToast(msg, duration = 2500) {
  const t = document.getElementById("toast");
  if (!t) return;
  t.textContent = msg;
  t.classList.add("show");
  setTimeout(() => t.classList.remove("show"), duration);
}

/** Escape HTML special characters */
export function escHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Format an API error object into a readable string */
export function formatError(data) {
  if (typeof data === "string") return data;
  return data?.error || data?.message || JSON.stringify(data);
}
