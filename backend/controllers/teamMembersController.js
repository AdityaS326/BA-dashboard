// backend/controllers/teamMembersController.js
// All team-member reads/writes go through a shared SQLite database so every
// user hitting the same backend instance sees the same data.

import db from "../utils/db.js";

const SEL_ALL = db.prepare(
  "SELECT name, role, status, leaves FROM team_members ORDER BY id"
);

const UPSERT = db.prepare(`
  INSERT INTO team_members (name, role, status, leaves, updated_at)
  VALUES (@name, @role, @status, @leaves, datetime('now'))
  ON CONFLICT(name) DO UPDATE SET
    role       = excluded.role,
    status     = excluded.status,
    leaves     = excluded.leaves,
    updated_at = datetime('now')
`);

function toRow(m) {
  return {
    name:   m.name   || "",
    role:   m.role   || "",
    status: m.status || "available",
    leaves: JSON.stringify(m.leaves || []),
  };
}

function fromRow(r) {
  return {
    name:   r.name,
    role:   r.role,
    status: r.status,
    leaves: (() => { try { return JSON.parse(r.leaves); } catch { return []; } })(),
  };
}

// GET /api/team-members
export function getMembers(_req, res) {
  try {
    res.json({ members: SEL_ALL.all().map(fromRow) });
  } catch (err) {
    console.error("[teamMembers] getMembers:", err.message);
    res.status(500).json({ error: err.message });
  }
}

// POST /api/team-members  { members: [...] }
// Upserts every member in the array; removes rows whose names are no longer present.
export function saveMembers(req, res) {
  const { members } = req.body;
  if (!Array.isArray(members))
    return res.status(400).json({ error: "members must be an array" });

  try {
    const syncAll = db.transaction((list) => {
      // Upsert each member
      for (const m of list) UPSERT.run(toRow(m));

      // Delete members that were removed from the list
      if (list.length > 0) {
        const placeholders = list.map(() => "?").join(",");
        db.prepare(
          `DELETE FROM team_members WHERE name NOT IN (${placeholders}) COLLATE NOCASE`
        ).run(...list.map(m => m.name));
      } else {
        db.prepare("DELETE FROM team_members").run();
      }
    });

    syncAll(members);
    res.json({ ok: true, members: SEL_ALL.all().map(fromRow) });
  } catch (err) {
    console.error("[teamMembers] saveMembers:", err.message);
    res.status(500).json({ error: err.message });
  }
}
