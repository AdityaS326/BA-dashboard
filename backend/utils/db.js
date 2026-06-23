// backend/utils/db.js
// Shared SQLite database — single instance reused across all requests.

import Database from "better-sqlite3";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import { mkdirSync, existsSync } from "fs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const DATA_DIR  = join(__dirname, "../data");
const DB_FILE   = join(DATA_DIR, "team_hub.db");

if (!existsSync(DATA_DIR)) mkdirSync(DATA_DIR, { recursive: true });

const db = new Database(DB_FILE);

// WAL mode → better concurrent reads (multiple users)
db.pragma("journal_mode = WAL");
db.pragma("foreign_keys = ON");

db.exec(`
  CREATE TABLE IF NOT EXISTS team_members (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    name       TEXT    NOT NULL UNIQUE COLLATE NOCASE,
    role       TEXT    NOT NULL DEFAULT '',
    status     TEXT    NOT NULL DEFAULT 'available',
    leaves     TEXT    NOT NULL DEFAULT '[]',
    updated_at TEXT    NOT NULL DEFAULT (datetime('now'))
  );
`);

// One-time migration: import legacy team-members.json if DB is still empty
import { readFileSync } from "fs";
const legacyFile = join(DATA_DIR, "team-members.json");
const isEmpty = db.prepare("SELECT COUNT(*) AS n FROM team_members").get().n === 0;
if (isEmpty && existsSync(legacyFile)) {
  try {
    const legacy = JSON.parse(readFileSync(legacyFile, "utf8"));
    const ins = db.prepare(`
      INSERT OR IGNORE INTO team_members (name, role, status, leaves)
      VALUES (@name, @role, @status, @leaves)
    `);
    db.transaction((rows) => {
      for (const m of rows) ins.run({
        name:   m.name   || "",
        role:   m.role   || "",
        status: m.status || "available",
        leaves: JSON.stringify(m.leaves || []),
      });
    })(legacy);
    console.log(`[db] Migrated ${legacy.length} member(s) from team-members.json → SQLite`);
  } catch (e) {
    console.warn("[db] Legacy migration skipped:", e.message);
  }
}

export default db;
