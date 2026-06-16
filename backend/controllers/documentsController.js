// backend/controllers/documentsController.js
// Uploads PDF / Word documents to SharePoint / OneDrive via Microsoft Graph API.

export async function uploadDocument(req, res) {
  const file = req.file;
  const { token, folderPath } = req.body;

  if (!file)  return res.status(400).json({ error: "No file provided." });
  if (!token) return res.status(400).json({ error: "No SharePoint token. Connect Microsoft 365 first." });

  const authHeader = token.startsWith("Bearer ") ? token : `Bearer ${token}`;
  const safeName   = file.originalname.replace(/[^a-zA-Z0-9._\- ()]/g, "_");

  try {
    let uploadUrl;
    const isShareUrl = /^https?:\/\//i.test((folderPath || "").trim());

    if (isShareUrl) {
      // Encode the sharing URL as a Graph sharing token: "u!" + base64url
      const b64 = Buffer.from(folderPath.trim()).toString("base64")
        .replace(/=+$/, "").replace(/\+/g, "-").replace(/\//g, "_");
      const sharingToken = "u!" + b64;

      const resolveResp = await fetch(
        `https://graph.microsoft.com/v1.0/shares/${sharingToken}/driveItem`,
        { headers: { Authorization: authHeader } }
      );
      if (!resolveResp.ok) {
        const err = await resolveResp.json().catch(() => ({}));
        return res.status(resolveResp.status).json({
          error: err.error?.message || "Could not access the SharePoint link — ensure you have edit access to this folder.",
        });
      }
      const item    = await resolveResp.json();
      const driveId = item.parentReference?.driveId;
      const itemId  = item.id;
      if (!driveId || !itemId)
        return res.status(400).json({ error: "Could not resolve drive location from sharing link." });

      uploadUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}:/${encodeURIComponent(safeName)}:/content`;
    } else {
      const folder      = (folderPath || "Documents/Falcon Dashboard").replace(/^\/|\/$/g, "");
      const encodedPath = `${folder}/${safeName}`.split("/").map(encodeURIComponent).join("/");
      uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodedPath}:/content`;
    }

    const resp = await fetch(uploadUrl, {
      method:  "PUT",
      headers: { Authorization: authHeader, "Content-Type": file.mimetype },
      body:    file.buffer,
    });

    if (!resp.ok) {
      const err = await resp.json().catch(() => ({}));
      return res.status(resp.status).json({ error: err.error?.message || `SharePoint upload failed (${resp.status})` });
    }

    const data = await resp.json();
    res.json({
      ok:   true,
      name: data.name,
      url:  data.webUrl,
      size: data.size,
      path: data.parentReference?.path || folderPath,
    });
  } catch (err) {
    res.status(502).json({ error: err.message });
  }
}
