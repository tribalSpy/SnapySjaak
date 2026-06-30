import pg from "pg";

const { Pool } = pg;

const connectionString = String(process.env.DATABASE_URL || "").trim();
const sslEnabled = ["1", "true", "yes", "on"].includes(String(process.env.DATABASE_SSL || "").trim().toLowerCase());

let pool = null;
let databaseStatus = {
  enabled: Boolean(connectionString),
  ready: false,
  error: connectionString ? "Not initialized yet" : "DATABASE_URL is not set",
};

function createPool() {
  if (!connectionString) {
    return null;
  }
  return new Pool({
    connectionString,
    ssl: sslEnabled ? { rejectUnauthorized: false } : false,
  });
}

export function isDatabaseEnabled() {
  return Boolean(connectionString);
}

export function getDatabaseStatus() {
  return { ...databaseStatus };
}

export async function dbQuery(text, params = []) {
  if (!pool) {
    throw new Error("Database is not initialized");
  }
  return pool.query(text, params);
}

export async function saveFustActionToDatabase(action) {
  if (!pool || !action?.id) {
    return;
  }

  await pool.query(
    `
      INSERT INTO fust_actions (
        id, type, action_date, week, day_name, country, customer_name, customer_code, connect_name,
        remark, fustbon_reference, fustfactuur_reference, dc, cctag, dcs, dco, pal, vk,
        deleted, created_by, created_at, updated_at
      )
      VALUES (
        $1, $2, $3, $4, $5, $6, $7, $8, $9,
        $10, $11, $12, $13, $14, $15, $16, $17, $18,
        $19, $20, COALESCE($21, now()), now()
      )
      ON CONFLICT (id) DO UPDATE SET
        type = EXCLUDED.type,
        action_date = EXCLUDED.action_date,
        week = EXCLUDED.week,
        day_name = EXCLUDED.day_name,
        country = EXCLUDED.country,
        customer_name = EXCLUDED.customer_name,
        customer_code = EXCLUDED.customer_code,
        connect_name = EXCLUDED.connect_name,
        remark = EXCLUDED.remark,
        fustbon_reference = EXCLUDED.fustbon_reference,
        fustfactuur_reference = EXCLUDED.fustfactuur_reference,
        dc = EXCLUDED.dc,
        cctag = EXCLUDED.cctag,
        dcs = EXCLUDED.dcs,
        dco = EXCLUDED.dco,
        pal = EXCLUDED.pal,
        vk = EXCLUDED.vk,
        deleted = EXCLUDED.deleted,
        created_by = EXCLUDED.created_by,
        created_at = COALESCE(fust_actions.created_at, EXCLUDED.created_at),
        updated_at = now()
    `,
    [
      action.id,
      action.type,
      action.action_date,
      action.week,
      action.day_name,
      action.country,
      action.customer_name,
      action.customer_code,
      action.connect_name,
      action.remark,
      action.fustbon_reference,
      action.fustfactuur_reference,
      Number(action.metrics?.dc || 0),
      Number(action.metrics?.cctag || 0),
      Number(action.metrics?.dcs || 0),
      Number(action.metrics?.dco || 0),
      Number(action.metrics?.pal || 0),
      Number(action.metrics?.vk || 0),
      action.deleted === true,
      action.created_by || "",
      action.created_at || null,
    ],
  );

  for (const documentKind of ["cmr", "fustbon"]) {
    const documentInfo = action?.[documentKind] || {};
    await pool.query(
      `
        INSERT INTO fust_action_documents (
          action_id, document_kind, status, file_id, file_name, web_link, mime_type, folder_id,
          error, uploaded_at, uploaded_by, updated_at
        )
        VALUES (
          $1, $2, $3, $4, $5, $6, $7, $8,
          $9, $10, $11, now()
        )
        ON CONFLICT (action_id, document_kind) DO UPDATE SET
          status = EXCLUDED.status,
          file_id = EXCLUDED.file_id,
          file_name = EXCLUDED.file_name,
          web_link = EXCLUDED.web_link,
          mime_type = EXCLUDED.mime_type,
          folder_id = EXCLUDED.folder_id,
          error = EXCLUDED.error,
          uploaded_at = EXCLUDED.uploaded_at,
          uploaded_by = EXCLUDED.uploaded_by,
          updated_at = now()
      `,
      [
        action.id,
        documentKind,
        documentInfo.status || "missing",
        documentInfo.file_id || "",
        documentInfo.file_name || "",
        documentInfo.web_link || "",
        documentInfo.mime_type || "",
        documentInfo.folder_id || "",
        documentInfo.error || "",
        documentInfo.uploaded_at || null,
        documentInfo.uploaded_by || "",
      ],
    );
  }
}

export async function markFustActionDeletedInDatabase(actionId) {
  if (!pool || !actionId) {
    return;
  }
  await pool.query(
    `
      UPDATE fust_actions
      SET deleted = true, updated_at = now()
      WHERE id = $1
    `,
    [actionId],
  );
}

export async function getFustDatabaseStats() {
  if (!pool) {
    return {
      total_actions: 0,
      active_actions: 0,
      deleted_actions: 0,
      document_rows: 0,
    };
  }

  const [actionsResult, documentsResult] = await Promise.all([
    pool.query(`
      SELECT
        COUNT(*)::int AS total_actions,
        COUNT(*) FILTER (WHERE deleted = false)::int AS active_actions,
        COUNT(*) FILTER (WHERE deleted = true)::int AS deleted_actions
      FROM fust_actions
    `),
    pool.query("SELECT COUNT(*)::int AS document_rows FROM fust_action_documents"),
  ]);

  return {
    total_actions: Number(actionsResult.rows?.[0]?.total_actions || 0),
    active_actions: Number(actionsResult.rows?.[0]?.active_actions || 0),
    deleted_actions: Number(actionsResult.rows?.[0]?.deleted_actions || 0),
    document_rows: Number(documentsResult.rows?.[0]?.document_rows || 0),
  };
}

const databaseMigrations = [
  `
    CREATE TABLE IF NOT EXISTS fust_actions (
      id uuid PRIMARY KEY,
      type text NOT NULL,
      action_date date NOT NULL,
      week integer,
      day_name text,
      country text NOT NULL,
      customer_name text NOT NULL,
      customer_code text,
      connect_name text,
      remark text,
      fustbon_reference text,
      fustfactuur_reference text,
      dc integer NOT NULL DEFAULT 0,
      cctag integer NOT NULL DEFAULT 0,
      dcs integer NOT NULL DEFAULT 0,
      dco integer NOT NULL DEFAULT 0,
      pal integer NOT NULL DEFAULT 0,
      vk integer NOT NULL DEFAULT 0,
      deleted boolean NOT NULL DEFAULT false,
      created_by text,
      created_at timestamptz NOT NULL DEFAULT now(),
      updated_at timestamptz NOT NULL DEFAULT now()
    )
  `,
  `
    CREATE TABLE IF NOT EXISTS fust_action_documents (
      id bigserial PRIMARY KEY,
      action_id uuid NOT NULL REFERENCES fust_actions(id) ON DELETE CASCADE,
      document_kind text NOT NULL,
      status text NOT NULL DEFAULT 'missing',
      file_id text,
      file_name text,
      web_link text,
      mime_type text,
      folder_id text,
      error text,
      uploaded_at timestamptz,
      uploaded_by text,
      created_at timestamptz NOT NULL DEFAULT now(),
      updated_at timestamptz NOT NULL DEFAULT now(),
      UNIQUE(action_id, document_kind)
    )
  `,
];

export async function initializeDatabase() {
  if (!connectionString) {
    databaseStatus = {
      enabled: false,
      ready: false,
      error: "DATABASE_URL is not set",
    };
    return databaseStatus;
  }

  try {
    pool = createPool();
    await pool.query("select 1");
    for (const migration of databaseMigrations) {
      await pool.query(migration);
    }
    databaseStatus = {
      enabled: true,
      ready: true,
      error: "",
    };
  } catch (error) {
    databaseStatus = {
      enabled: true,
      ready: false,
      error: error instanceof Error ? error.message : String(error),
    };
    throw error;
  }

  return getDatabaseStatus();
}
