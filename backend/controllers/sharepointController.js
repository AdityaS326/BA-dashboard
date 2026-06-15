// backend/controllers/sharepointController.js
import { config } from "../config/index.js";
import { buildReportDocx } from "../utils/docxBuilder.js";

// â"€â"€ Test Graph API connection â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
export async function testConnection(req, res) {
  const { token } = req.body;
  if (!token) return res.status(400).json({ ok: false, error: "No access token provided." });

  const bearer = token.startsWith("Bearer ") ? token : `Bearer ${token}`;

  const [userRes, driveRes] = await Promise.all([
    fetch("https://graph.microsoft.com/v1.0/me",       { headers: { Authorization: bearer } }),
    fetch("https://graph.microsoft.com/v1.0/me/drive", { headers: { Authorization: bearer } }),
  ]);

  if (!userRes.ok) {
    const err = await userRes.json();
    return res.json({ ok: false, error: err.error?.message || "Invalid or expired token." });
  }

  const user  = await userRes.json();
  const drive = driveRes.ok ? await driveRes.json() : null;

  res.json({
    ok:       true,
    user:     user.displayName,
    email:    user.mail || user.userPrincipalName,
    driveId:  drive?.id,
    driveType: drive?.driveType,
  });
}

// â"€â"€ Export report as .docx â†' upload to SharePoint â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
export async function exportReport(req, res) {
  const {
    token,
    spUrl,
    filename = "Weekly_Report.docx",
    reportContent,
    reportMeta = {},
  } = req.body;

  if (!token)         return res.status(400).json({ ok: false, error: "No access token." });
  if (!reportContent) return res.status(400).json({ ok: false, error: "No report content. Generate report first." });

  const bearer = token.startsWith("Bearer ") ? token : `Bearer ${token}`;
  const { name = "Aditya S", dept = "Technology", week = new Date().toLocaleDateString("en-IN"), manager = "Igor" } = reportMeta;

  // Build .docx
  const docBuffer = await buildReportDocx({ reportText: reportContent, name, dept, week, manager });

  // Parse SharePoint folder path from URL
  let uploadPath = "/Weekly Reports";
  try {
    const urlObj = new URL(spUrl);
    const idParam = urlObj.searchParams.get("id");
    if (idParam) {
      const decoded = decodeURIComponent(idParam);
      const docsIdx = decoded.indexOf("/Documents/");
      uploadPath =
        docsIdx !== -1
          ? decoded.substring(docsIdx + "/Documents".length)
          : "/" + decoded.split("/").slice(-3).join("/");
    }
  } catch (_) {}

  // Upload via Graph API
  const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:${uploadPath}/${filename}:/content`;
  const uploadRes = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      Authorization:  bearer,
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    },
    body: docBuffer,
  });

  if (!uploadRes.ok) {
    const uploadErr = await uploadRes.json();
    // Fallback: try /Weekly Reports/ at root
    if (uploadErr.error?.code === "itemNotFound") {
      const fbUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/Weekly%20Reports/${filename}:/content`;
      const fbRes  = await fetch(fbUrl, { method: "PUT", headers: { Authorization: bearer, "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }, body: docBuffer });
      if (!fbRes.ok) { const e = await fbRes.json(); return res.json({ ok: false, error: e.error?.message }); }
      const fb = await fbRes.json();
      return res.json({ ok: true, fileId: fb.id, webUrl: fb.webUrl, path: "/Weekly Reports/" + filename, note: "Saved to /Weekly Reports/" });
    }
    return res.json({ ok: false, error: uploadErr.error?.message || JSON.stringify(uploadErr) });
  }

  const ud = await uploadRes.json();
  res.json({ ok: true, fileId: ud.id, webUrl: ud.webUrl, path: (ud.parentReference?.path || "") + "/" + filename });
}



// â"€â"€ OAuth callback â€" exchange code for token â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
export async function listFiles(req, res) {
  let tok = req.query.token;
  if (!tok) tok = req.headers.authorization;
  if (!tok && req.body) tok = req.body.token;
  if (!tok) return res.status(400).json({ error: 'No access token.' });
  tok = String(tok).replace(/^Bearer\s+/i, '').trim();
  if (!tok) return res.status(400).json({ error: 'No access token.' });
  const bearer = 'Bearer ' + tok;
  const url = 'https://graph.microsoft.com/v1.0/me/drive/recent?$select=id,name,size,lastModifiedDateTime,webUrl,file,parentReference&$top=100';
  try {
    const r = await fetch(url, { headers: { Authorization: bearer } });
    if (!r.ok) {
      const e = await r.json().catch(function() { return {}; });
      return res.json({ error: (e.error && e.error.message) || ('Graph error ' + r.status) });
    }
    const data = await r.json();
    const docExts = /\.(pdf|doc|docx|xlsx|xls|pptx|ppt|txt|md|csv|odt|rtf)$/i;
    const items = data.value || [];
    const files = [];
    for (let i = 0; i < items.length; i++) {
      const f = items[i];
      if (!f.file || !docExts.test(f.name)) continue;
      const ref = f.parentReference || {};
      files.push({ id: f.id, name: f.name, size: f.size, modified: f.lastModifiedDateTime, webUrl: f.webUrl, folder: ref.name || '', path: ref.path || '' });
    }
    return res.json({ files: files });
  } catch (err) {
    return res.status(502).json({ error: err.message });
  }
}

export async function listSites(req, res) {
  let tok = req.query.token;
  if (!tok) tok = req.headers.authorization;
  if (!tok) return res.status(400).json({ error: 'No access token.' });
  tok = String(tok).replace(/^Bearer\s+/i, '').trim();
  const bearer = 'Bearer ' + tok;
  const graph = 'https://graph.microsoft.com/v1.0';

  function normSites(arr) {
    return arr.map(function(s) { return { id: s.id, name: s.displayName || s.name || s.webUrl, webUrl: s.webUrl }; });
  }

  try {
    // Attempt 1: search all sites (requires Sites.Read.All)
    const r1 = await fetch(graph + '/sites?search=*&$select=id,displayName,webUrl,name&$top=50', { headers: { Authorization: bearer } });
    if (r1.ok) {
      const d1 = await r1.json();
      if (d1.value && d1.value.length > 0) return res.json({ sites: normSites(d1.value), source: 'search' });
    }

    // Attempt 2: followed sites (no admin consent needed)
    const r2 = await fetch(graph + '/me/followedSites?$select=id,displayName,webUrl,name&$top=50', { headers: { Authorization: bearer } });
    if (r2.ok) {
      const d2 = await r2.json();
      if (d2.value && d2.value.length > 0) return res.json({ sites: normSites(d2.value), source: 'followed' });
    }

    // Attempt 3: root SharePoint site only
    const r3 = await fetch(graph + '/sites/root?$select=id,displayName,webUrl,name', { headers: { Authorization: bearer } });
    if (r3.ok) {
      const d3 = await r3.json();
      return res.json({ sites: [{ id: d3.id, name: d3.displayName || 'Root site', webUrl: d3.webUrl }], source: 'root' });
    }

    const eBody = await r3.json().catch(function() { return {}; });
    return res.json({ error: (eBody.error && eBody.error.message) || 'Could not load sites. Check token permissions (Sites.Read.All or Sites.ReadWrite.All).' });
  } catch (err) {
    return res.status(502).json({ error: err.message });
  }
}

export async function listSiteFiles(req, res) {
  let tok = req.query.token;
  if (!tok) tok = req.headers.authorization;
  if (!tok) return res.status(400).json({ error: 'No access token.' });
  tok = String(tok).replace(/^Bearer\s+/i, '').trim();
  const siteId = req.query.siteId;
  if (!siteId) return res.status(400).json({ error: 'siteId is required.' });
  const bearer = 'Bearer ' + tok;
  const graph = 'https://graph.microsoft.com/v1.0';
  try {
    // Step 1: get all lists in the site
    const listsUrl = graph + '/sites/' + siteId + '/lists?$select=id,displayName,list&$top=50';
    const lr = await fetch(listsUrl, { headers: { Authorization: bearer } });
    if (!lr.ok) {
      const e = await lr.json().catch(function() { return {}; });
      return res.json({ error: (e.error && e.error.message) || ('Lists error ' + lr.status) });
    }
    const listsData = await lr.json();
    // Keep only document libraries (template = documentLibrary)
    const docLibs = (listsData.value || []).filter(function(l) {
      return l.list && l.list.template === 'documentLibrary';
    });

    // Step 2: for each doc library, fetch items with driveItem expanded
    const docExts = /\.(pdf|doc|docx|xlsx|xls|pptx|ppt|txt|md|csv|odt|rtf)$/i;
    const files = [];
    for (let i = 0; i < docLibs.length; i++) {
      const lib = docLibs[i];
      const itemsUrl = graph + '/sites/' + siteId + '/lists/' + lib.id + '/items?$expand=driveItem($select=id,name,webUrl,lastModifiedDateTime,file,size)&$top=100';
      const ir = await fetch(itemsUrl, { headers: { Authorization: bearer } });
      if (!ir.ok) continue;
      const itemsData = await ir.json();
      const items = itemsData.value || [];
      for (let j = 0; j < items.length; j++) {
        const di = items[j].driveItem;
        if (!di || !di.file || !docExts.test(di.name)) continue;
        files.push({
          id: di.id,
          name: di.name,
          size: di.size,
          modified: di.lastModifiedDateTime,
          webUrl: di.webUrl,
          folder: lib.displayName,
          path: lib.displayName,
        });
      }
    }
    return res.json({ files: files, siteId: siteId, libCount: docLibs.length });
  } catch (err) {
    return res.status(502).json({ error: err.message });
  }
}
export async function oauthCallback(req, res) {
  const { code, error, error_description } = req.query;
  if (error) return res.redirect(`/?sp_error=${encodeURIComponent(error_description || error)}`);
  if (!code)  return res.status(400).json({ error: "No authorization code received." });

  const { tenantId, clientId, clientSecret } = config.sharepoint;
  const origin      = process.env.VERCEL_URL ? `https://${process.env.VERCEL_URL}` : `http://localhost:${config.port}`;
  const redirectUri = `${origin}/api/sharepoint/callback`;

  if (!tenantId || !clientId || !clientSecret)
    return res.status(500).json({ error: "SP_TENANT_ID / SP_CLIENT_ID / SP_CLIENT_SECRET not set." });

  const tokenRes = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ grant_type: "authorization_code", client_id: clientId, client_secret: clientSecret, code, redirect_uri: redirectUri, scope: "Files.ReadWrite.All Sites.ReadWrite.All User.Read offline_access" }),
    }
  );

  const tokenData = await tokenRes.json();
  if (tokenData.error) return res.redirect(`/?sp_error=${encodeURIComponent(tokenData.error_description || tokenData.error)}`);

  const rt = tokenData.refresh_token ? `&sp_refresh=${encodeURIComponent(tokenData.refresh_token)}` : "";
  res.redirect(`/?sp_token=${encodeURIComponent(tokenData.access_token)}${rt}`);
}

// POST /api/sharepoint/refresh — exchange a refresh_token for a new access_token
export async function refreshToken(req, res) {
  const { refreshToken: rt } = req.body;
  if (!rt) return res.status(400).json({ error: "refreshToken is required." });

  const { tenantId, clientId, clientSecret } = config.sharepoint;
  if (!tenantId || !clientId || !clientSecret)
    return res.status(500).json({ error: "OAuth not configured. Set SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET in .env" });

  try {
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      {
        method:  "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          grant_type:    "refresh_token",
          client_id:     clientId,
          client_secret: clientSecret,
          refresh_token: rt,
          scope: "Files.ReadWrite.All Sites.ReadWrite.All User.Read offline_access Calendars.Read Mail.Read Mail.Send",
        }),
      }
    );
    const data = await tokenRes.json();
    if (data.error) return res.status(401).json({ error: data.error_description || data.error });
    res.json({
      accessToken:  data.access_token,
      refreshToken: data.refresh_token || rt,
      expiresIn:    data.expires_in,
    });
  } catch (err) {
    res.status(502).json({ error: err.message });
  }
}




