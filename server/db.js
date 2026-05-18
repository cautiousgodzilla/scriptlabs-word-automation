import Database from "better-sqlite3";
import { mkdirSync, existsSync } from "fs";
import { join } from "path";
import { homedir } from "os";
import { BUILTIN_WORKFLOWS } from "./workflows/builtins.js";

const DB_DIR = join(homedir(), ".office-agent");
const DB_PATH = join(DB_DIR, "workflows.db");

if (!existsSync(DB_DIR)) mkdirSync(DB_DIR, { recursive: true });

export const db = new Database(DB_PATH);

db.exec(`
  CREATE TABLE IF NOT EXISTS workflows (
    id             TEXT PRIMARY KEY,
    title          TEXT NOT NULL,
    type           TEXT NOT NULL,
    prompt_md      TEXT,
    columns_config TEXT,
    practice       TEXT,
    is_system      INTEGER NOT NULL DEFAULT 0,
    created_at     TEXT NOT NULL DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS hidden_workflows (
    workflow_id TEXT PRIMARY KEY
  );
`);

const upsertBuiltin = db.prepare(`
  INSERT INTO workflows (id, title, type, prompt_md, practice, is_system)
  VALUES (@id, @title, @type, @prompt_md, @practice, 1)
  ON CONFLICT(id) DO UPDATE SET
    title     = excluded.title,
    prompt_md = excluded.prompt_md,
    practice  = excluded.practice
`);

db.transaction(() => {
  for (const w of BUILTIN_WORKFLOWS) {
    upsertBuiltin.run({
      id: w.id,
      title: w.title,
      type: w.type,
      prompt_md: w.prompt_md,
      practice: w.practice ?? null,
    });
  }
})();

// Built fresh per-message — cheap read, no invalidation needed.
export function buildWorkflowStore() {
  const store = new Map();

  // Built-ins first so user workflows with the same id override
  for (const w of BUILTIN_WORKFLOWS) {
    if (w.prompt_md) store.set(w.id, { title: w.title, prompt_md: w.prompt_md });
  }

  const rows = db
    .prepare(`SELECT id, title, prompt_md FROM workflows WHERE type='assistant' AND prompt_md IS NOT NULL AND is_system=0`)
    .all();
  for (const r of rows) store.set(r.id, { title: r.title, prompt_md: r.prompt_md });

  return store;
}
