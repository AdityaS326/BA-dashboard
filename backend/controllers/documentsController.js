// backend/controllers/documentsController.js
// Uploads PDF / Word documents to SharePoint / OneDrive via Microsoft Graph API.

export async function uploadDocument(req, res) {
  const file       = req.file;
  const { token, folderPath } = req.body;

  if (!file)  return res.status(400).json({ error: "No file provided." });
  if (!token) return res.status(400).json({ error: "No SharePoint token. Login with Microsoft 365 first (Weekly report → OAuth setup)." });

  const safeName = file.originalname.replace(/[^a-zA-Z0-9._\- ()]/g, "_");
  const folder   = (folderPath || "Documents/Falcon Dashboard").replace(/^\/|\/$/g, "");
  const path     = `${folder}/${safeName}`;

  try {
    const encodedPath = path.split("/").map(encodeURIComponent).join("/");
    const resp = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/root:/${encodedPath}:/content`,
      {
        method:  "PUT",
        headers: {
          "Authorization": token.startsWith("Bearer ") ? token : `Bearer ${token}`,
          "Content-Type":  file.mimetype,
        },
        body: file.buffer,
      }
    );

    if (!resp.ok) {
      const err = await resp.json().catch(() => ({}));
      return res.status(resp.status).json({ error: err.error?.message || `SharePoint upload failed (${resp.status})` });
    }

    const data = await resp.json();
    res.json({
      ok:       true,
      name:     data.name,
      url:      data.webUrl,
      size:     data.size,
      path:     data.parentReference?.path || folder,
    });
  } catch (err) {
    res.status(502).json({ error: err.message });
  }
}
