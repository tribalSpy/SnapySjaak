// db.js — SQLite layer. better-sqlite3 is synchroon, dus geen async/await nodig.
const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');

function createDb(dbPath) {
  const dir = path.dirname(dbPath);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

  const db = new Database(dbPath);
  db.pragma('journal_mode = WAL');
  db.pragma('foreign_keys = ON');

  // Schema
  db.exec(`
    CREATE TABLE IF NOT EXISTS articles (
      broncode          INTEGER PRIMARY KEY,
      omschrijving      TEXT NOT NULL,
      variant_code      TEXT NOT NULL,
      broncode_inlezen  INTEGER,
      afronden          INTEGER,
      active            INTEGER NOT NULL DEFAULT 1,
      updated_at        TEXT NOT NULL DEFAULT (datetime('now'))
    );

    CREATE INDEX IF NOT EXISTS idx_articles_omschrijving ON articles(omschrijving);
    CREATE INDEX IF NOT EXISTS idx_articles_variant ON articles(variant_code);

    CREATE TABLE IF NOT EXISTS ape_lookup (
      omschrijving      TEXT PRIMARY KEY,
      ape               INTEGER NOT NULL,
      updated_at        TEXT NOT NULL DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS soortblad_overrides (
      original_name     TEXT PRIMARY KEY,
      display_name      TEXT,
      exclude           INTEGER NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS import_runs (
      id                INTEGER PRIMARY KEY AUTOINCREMENT,
      created_at        TEXT NOT NULL DEFAULT (datetime('now')),
      user              TEXT,
      row_count         INTEGER,
      total_som         INTEGER,
      status            TEXT,
      warnings_json     TEXT,
      raw_json          TEXT
    );
  `);

  // Migraties: voeg kolommen toe die er in oudere DBs nog niet zijn
  const runCols = new Set(db.prepare('PRAGMA table_info(import_runs)').all().map(c => c.name));
  if (!runCols.has('result_json')) {
    db.exec('ALTER TABLE import_runs ADD COLUMN result_json TEXT');
  }
  if (!runCols.has('vertrek_datum')) {
    db.exec('ALTER TABLE import_runs ADD COLUMN vertrek_datum TEXT');
  }
  if (!runCols.has('label')) {
    db.exec('ALTER TABLE import_runs ADD COLUMN label TEXT');
  }

  const articleCols = new Set(db.prepare('PRAGMA table_info(articles)').all().map(c => c.name));
  if (!articleCols.has('zonder_tak')) {
    db.exec('ALTER TABLE articles ADD COLUMN zonder_tak INTEGER NOT NULL DEFAULT 0');
  }

  // Prepared statements
  const stmts = {
    // Articles
    getArticle: db.prepare('SELECT * FROM articles WHERE broncode = ?'),
    listArticles: db.prepare('SELECT * FROM articles WHERE active = 1 ORDER BY omschrijving'),
    upsertArticle: db.prepare(`
      INSERT INTO articles (broncode, omschrijving, variant_code, broncode_inlezen, afronden, zonder_tak, active, updated_at)
      VALUES (@broncode, @omschrijving, @variant_code, @broncode_inlezen, @afronden, @zonder_tak, 1, datetime('now'))
      ON CONFLICT(broncode) DO UPDATE SET
        omschrijving = excluded.omschrijving,
        variant_code = excluded.variant_code,
        broncode_inlezen = excluded.broncode_inlezen,
        afronden = excluded.afronden,
        zonder_tak = excluded.zonder_tak,
        updated_at = datetime('now')
    `),
    deactivateArticle: db.prepare('UPDATE articles SET active = 0, updated_at = datetime(\'now\') WHERE broncode = ?'),

    // APE
    getApe: db.prepare('SELECT ape FROM ape_lookup WHERE omschrijving = ?'),
    listApe: db.prepare('SELECT * FROM ape_lookup ORDER BY omschrijving'),
    upsertApe: db.prepare(`
      INSERT INTO ape_lookup (omschrijving, ape, updated_at)
      VALUES (@omschrijving, @ape, datetime('now'))
      ON CONFLICT(omschrijving) DO UPDATE SET
        ape = excluded.ape,
        updated_at = datetime('now')
    `),
    deleteApe: db.prepare('DELETE FROM ape_lookup WHERE omschrijving = ?'),

    // Soortblad overrides
    listOverrides: db.prepare('SELECT * FROM soortblad_overrides'),
    upsertOverride: db.prepare(`
      INSERT INTO soortblad_overrides (original_name, display_name, exclude)
      VALUES (@original_name, @display_name, @exclude)
      ON CONFLICT(original_name) DO UPDATE SET
        display_name = excluded.display_name,
        exclude = excluded.exclude
    `),

    // Import audit
    logRun: db.prepare(`
      INSERT INTO import_runs (user, row_count, total_som, status, warnings_json, raw_json, result_json, vertrek_datum, label)
      VALUES (@user, @row_count, @total_som, @status, @warnings_json, @raw_json, @result_json, @vertrek_datum, @label)
    `),
    listRuns: db.prepare('SELECT id, created_at, user, row_count, total_som, status, vertrek_datum, label FROM import_runs ORDER BY id DESC LIMIT 100'),
    getRun: db.prepare('SELECT * FROM import_runs WHERE id = ?'),
    updateRunDate: db.prepare('UPDATE import_runs SET vertrek_datum = ?, result_json = ? WHERE id = ?'),
    updateRunLabel: db.prepare('UPDATE import_runs SET label = ? WHERE id = ?'),
    deleteRun: db.prepare('DELETE FROM import_runs WHERE id = ?'),
  };

  // Convenience helpers
  return {
    db,
    getArticleByBroncode(broncode) {
      return stmts.getArticle.get(broncode);
    },
    listArticles() {
      return stmts.listArticles.all();
    },
    upsertArticle(a) {
      return stmts.upsertArticle.run({
        broncode: a.broncode,
        omschrijving: a.omschrijving,
        variant_code: a.variant_code,
        broncode_inlezen: a.broncode_inlezen ?? null,
        afronden: a.afronden ?? null,
        zonder_tak: a.zonder_tak ? 1 : 0,
      });
    },
    deactivateArticle(broncode) {
      return stmts.deactivateArticle.run(broncode);
    },
    getApe(omschrijving) {
      const row = stmts.getApe.get(omschrijving);
      return row ? row.ape : null;
    },
    listApe() {
      return stmts.listApe.all();
    },
    upsertApe(a) {
      return stmts.upsertApe.run({
        omschrijving: a.omschrijving,
        ape: a.ape,
      });
    },
    deleteApe(omschrijving) {
      return stmts.deleteApe.run(omschrijving);
    },
    listOverrides() {
      return stmts.listOverrides.all();
    },
    upsertOverride(o) {
      return stmts.upsertOverride.run({
        original_name: o.original_name,
        display_name: o.display_name ?? null,
        exclude: o.exclude ? 1 : 0,
      });
    },
    logRun(run) {
      return stmts.logRun.run({
        user: run.user ?? null,
        row_count: run.row_count ?? null,
        total_som: run.total_som ?? null,
        status: run.status ?? null,
        warnings_json: run.warnings_json ?? null,
        raw_json: run.raw_json ?? null,
        result_json: run.result_json ?? null,
        vertrek_datum: run.vertrek_datum ?? null,
        label: run.label ?? null,
      });
    },
    listRuns() {
      return stmts.listRuns.all();
    },
    getRun(id) {
      return stmts.getRun.get(id);
    },
    updateRunDate(id, vertrekDatum, resultJson) {
      return stmts.updateRunDate.run(vertrekDatum, resultJson, id);
    },
    updateRunLabel(id, label) {
      return stmts.updateRunLabel.run(label, id);
    },
    deleteRun(id) {
      return stmts.deleteRun.run(id);
    },
    // Bulk import via transaction
    bulkUpsertArticles(articles) {
      const tx = db.transaction((items) => {
        for (const a of items) {
          stmts.upsertArticle.run({
            broncode: a.broncode,
            omschrijving: a.omschrijving,
            variant_code: a.variant_code,
            broncode_inlezen: a.broncode_inlezen ?? null,
            afronden: a.afronden ?? null,
            zonder_tak: a.zonder_tak ? 1 : 0,
          });
        }
      });
      tx(articles);
    },
    bulkUpsertApe(entries) {
      const tx = db.transaction((items) => {
        for (const e of items) {
          stmts.upsertApe.run({ omschrijving: e.omschrijving, ape: e.ape });
        }
      });
      tx(entries);
    },
  };
}

module.exports = { createDb };
