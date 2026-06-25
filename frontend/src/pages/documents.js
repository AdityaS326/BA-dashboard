// frontend/src/pages/documents.js
// Documents are user-managed and persisted to localStorage — no hardcoded data.

function loadDocs() {
  try {
    const raw = localStorage.getItem("ba_docs");
    return raw ? JSON.parse(raw) : [];
  } catch { return []; }
}

function saveDocs() {
  try { localStorage.setItem("ba_docs", JSON.stringify(DOCS)); } catch {}
}

export let DOCS = loadDocs();

const BADGE_MAP = { Approved: "b-green", Pending: "b-amber", Draft: "b-amber", Delivered: "b-green", Overdue: "b-red" };

let _activeMenu = null;

function closeActiveMenu() {
  if (_activeMenu) { _activeMenu.remove(); _activeMenu = null; }
}

function showDocMenu(e, idx) {
  e.stopPropagation();
  closeActiveMenu();

  const menu = document.createElement("div");
  menu.style.cssText = "position:fixed;z-index:9999;background:var(--surface);border:1px solid var(--border);border-radius:var(--r-sm);box-shadow:0 4px 16px rgba(0,0,0,.18);min-width:150px;overflow:hidden";

  [
    { icon: "ti-edit",  label: "Edit name", fn: () => editDoc(idx) },
    { icon: "ti-trash", label: "Delete",    fn: () => deleteDoc(idx), red: true },
  ].forEach(({ icon, label, fn, red }) => {
    const btn = document.createElement("button");
    btn.style.cssText = `display:flex;align-items:center;gap:9px;width:100%;padding:9px 14px;background:none;border:none;cursor:pointer;font-size:13px;color:${red ? "var(--red)" : "var(--text)"};text-align:left`;
    btn.innerHTML = `<i class="ti ${icon}" style="font-size:13px;color:${red ? "var(--red)" : "var(--muted)"}"></i>${label}`;
    btn.onmouseenter = () => { btn.style.background = "var(--surface2)"; };
    btn.onmouseleave = () => { btn.style.background = "none"; };
    btn.onclick = () => { closeActiveMenu(); fn(); };
    menu.appendChild(btn);
  });

  const rect = e.target.getBoundingClientRect();
  menu.style.top  = (rect.bottom + 4) + "px";
  menu.style.left = Math.max(8, rect.right - 154) + "px";
  document.body.appendChild(menu);
  _activeMenu = menu;
  setTimeout(() => document.addEventListener("click", closeActiveMenu, { once: true }), 10);
}

function editDoc(idx) {
  const d    = DOCS[idx];
  const name = prompt("Edit document name:", d.n);
  if (name && name.trim()) { DOCS[idx].n = name.trim(); saveDocs(); filterDocs(); }
}

function deleteDoc(idx) {
  if (!confirm(`Delete "${DOCS[idx].n}"?`)) return;
  DOCS.splice(idx, 1);
  saveDocs();
  filterDocs();
  const det = document.getElementById("doc-detail");
  if (det) det.style.display = "none";
}

function triggerDocUpload(idx) {
  const input = document.createElement("input");
  input.type   = "file";
  input.accept = ".pdf,.doc,.docx";
  input.onchange = () => {
    if (!input.files[0]) return;
    let path = localStorage.getItem("spDocFolder");
    if (!path) {
      path = prompt("SharePoint / OneDrive folder path for upload:", "Documents/BA Hub");
      if (!path) return;
      localStorage.setItem("spDocFolder", path);
    }
    window._uploadDocFile && window._uploadDocFile(idx, input.files[0], path);
  };
  input.click();
}

export function addNewDoc() {
  const name = prompt("Document name:", "New Document");
  if (!name) return;
  const today = new Date().toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
  DOCS.unshift({ n: name.trim(), v: "v1.0", s: "Draft", d: today, desc: "Newly added.", url: "" });
  saveDocs();
  filterDocs();
}

export function renderDocs(list) {
  const tb = document.getElementById("doc-tbody");
  if (!tb) return;
  tb.innerHTML = "";
  list.forEach((d) => {
    const realIdx = DOCS.indexOf(d);
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td title="${d.n}" style="cursor:pointer" onclick="window._showDocDetail(${realIdx})">
        ${d.n}${d.url ? ` <a href="${d.url}" target="_blank" style="color:var(--blue);font-size:10px;margin-left:3px" onclick="event.stopPropagation()"><i class="ti ti-external-link"></i></a>` : ""}
      </td>
      <td>${d.v}</td>
      <td><span class="badge ${BADGE_MAP[d.s] || "b-gray"}">${d.s}</span></td>
      <td>
        <div style="display:flex;align-items:center;justify-content:space-between;gap:4px">
          <span>${d.d}</span>
          <button onclick="event.stopPropagation();window.showDocMenu(event,${realIdx})"
            style="background:none;border:none;cursor:pointer;padding:2px 5px;color:var(--muted);font-size:18px;line-height:1;border-radius:4px;flex-shrink:0"
            title="Actions">⋯</button>
        </div>
      </td>`;
    tb.appendChild(tr);
  });
}

export function filterDocs() {
  const f = document.getElementById("doc-filter")?.value || "";
  const q = (document.getElementById("doc-search")?.value || "").toLowerCase();
  renderDocs(DOCS.filter((d) => (!f || d.s === f) && (!q || d.n.toLowerCase().includes(q))));
  // Update doc chart and stat card whenever docs change
  window._refreshDocStats?.();
}

window.showDocMenu = showDocMenu;

window._showDocDetail = function (realIdx) {
  const d    = DOCS[realIdx];
  if (!d) return;
  const det  = document.getElementById("doc-detail");
  const name = document.getElementById("doc-detail-name");
  const body = document.getElementById("doc-detail-body");
  if (!det || !name || !body) return;
  det.style.display = "block";
  name.textContent  = d.n;
  body.innerHTML    = `
    <div style="margin-bottom:8px">
      <span class="badge ${BADGE_MAP[d.s] || "b-gray"}" style="margin-right:8px">${d.s}</span>
      <span style="color:var(--muted)">${d.v} · ${d.d}</span>
    </div>
    <div style="margin-bottom:12px;font-size:13px;color:var(--muted)">${d.desc}</div>
    <div style="margin-bottom:12px">
      <label style="font-size:11px;font-weight:500;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:5px;display:block">Source link</label>
      <div style="display:flex;gap:6px;align-items:center">
        <input id="doc-url-input" value="${d.url || ""}" placeholder="Paste SharePoint or any link..." style="flex:1;font-size:12px;font-family:var(--mono);margin:0">
        <button class="sm primary" onclick="window.setDocUrl(${realIdx})"><i class="ti ti-device-floppy" style="font-size:11px"></i> Save</button>
        ${d.url ? `<a href="${d.url}" target="_blank" onclick="event.stopPropagation()"><button class="sm"><i class="ti ti-external-link" style="font-size:11px"></i></button></a>` : ""}
      </div>
    </div>
    <div style="display:flex;gap:6px;flex-wrap:wrap">
      <button class="sm" onclick="window.triggerDocUploadFromDetail(${realIdx})"><i class="ti ti-cloud-upload" style="font-size:12px"></i> Upload</button>
      <button class="sm" onclick="window.editDocFromDetail(${realIdx})"><i class="ti ti-edit" style="font-size:12px"></i> Edit</button>
      <button class="sm" style="color:var(--red)" onclick="window.deleteDocFromDetail(${realIdx})"><i class="ti ti-trash" style="font-size:12px"></i> Delete</button>
    </div>`;
};

window.triggerDocUploadFromDetail = triggerDocUpload;
window.editDocFromDetail          = editDoc;
window.deleteDocFromDetail        = deleteDoc;

window.setDocUrl = function (realIdx) {
  const input = document.getElementById("doc-url-input");
  if (!input) return;
  const url = input.value.trim();
  if (realIdx >= 0 && realIdx < DOCS.length) {
    DOCS[realIdx].url = url;
    saveDocs();
    filterDocs();
    window._showDocDetail(realIdx);
  }
};
