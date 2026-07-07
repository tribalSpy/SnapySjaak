import pg from "pg";
import crypto from "node:crypto";

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
        deleted, created_by, created_at, confirmed_at, confirmed_by, import_source, confirmation_reminder, updated_at
      )
      VALUES (
        $1, $2, $3, $4, $5, $6, $7, $8, $9,
        $10, $11, $12, $13, $14, $15, $16, $17, $18,
        $19, $20, COALESCE($21, now()), $22, $23, $24::jsonb, $25::jsonb, now()
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
        confirmed_at = EXCLUDED.confirmed_at,
        confirmed_by = EXCLUDED.confirmed_by,
        import_source = EXCLUDED.import_source,
        confirmation_reminder = EXCLUDED.confirmation_reminder,
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
      action.confirmed_at || null,
      action.confirmed_by || "",
      JSON.stringify(action.import_source || {}),
      JSON.stringify(action.confirmation_reminder || {}),
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

function jsonValue(value, fallback = {}) {
  if (value && typeof value === "object") {
    return value;
  }
  if (typeof value === "string") {
    try {
      return JSON.parse(value);
    } catch {
      return fallback;
    }
  }
  return fallback;
}

function hashLlmAgentKey(apiKey) {
  return crypto.createHash("sha256").update(String(apiKey || "")).digest("hex");
}

function mapLlmAgentRow(row) {
  return {
    agent_name: String(row?.agent_name || "").trim(),
    pc_name: String(row?.pc_name || "").trim(),
    model_name: String(row?.model_name || "").trim(),
    version: String(row?.version || "").trim(),
    status: String(row?.status || "").trim() || "offline",
    capabilities: jsonValue(row?.capabilities, []),
    meta: jsonValue(row?.meta, {}),
    last_seen_at: row?.last_seen_at ? new Date(row.last_seen_at).toISOString() : "",
    last_job_claimed_at: row?.last_job_claimed_at ? new Date(row.last_job_claimed_at).toISOString() : "",
    created_at: row?.created_at ? new Date(row.created_at).toISOString() : "",
    updated_at: row?.updated_at ? new Date(row.updated_at).toISOString() : "",
  };
}

function mapLlmJobRow(row) {
  return {
    id: String(row?.id || "").trim(),
    job_type: String(row?.job_type || "").trim(),
    status: String(row?.status || "").trim() || "pending",
    created_by: String(row?.created_by || "").trim(),
    shipment_id: String(row?.shipment_id || "").trim(),
    collection_id: String(row?.collection_id || "").trim(),
    document_kind: String(row?.document_kind || "").trim(),
    priority: Number(row?.priority || 0),
    attempt_count: Number(row?.attempt_count || 0),
    max_attempts: Number(row?.max_attempts || 3),
    agent_name: String(row?.agent_name || "").trim(),
    payload_json: jsonValue(row?.payload_json, {}),
    result_json: jsonValue(row?.result_json, {}),
    error_text: String(row?.error_text || "").trim(),
    created_at: row?.created_at ? new Date(row.created_at).toISOString() : "",
    claimed_at: row?.claimed_at ? new Date(row.claimed_at).toISOString() : "",
    finished_at: row?.finished_at ? new Date(row.finished_at).toISOString() : "",
    updated_at: row?.updated_at ? new Date(row.updated_at).toISOString() : "",
  };
}

export async function upsertLlmAgentHeartbeat(agent, apiKey) {
  if (!pool || !String(agent?.agent_name || "").trim() || !String(apiKey || "").trim()) {
    return null;
  }
  const result = await pool.query(
    `
      INSERT INTO llm_agents (
        agent_name, api_key_hash, pc_name, model_name, version, status, capabilities, meta,
        last_seen_at, updated_at
      )
      VALUES (
        $1, $2, $3, $4, $5, $6, $7::jsonb, $8::jsonb,
        now(), now()
      )
      ON CONFLICT (agent_name) DO UPDATE SET
        api_key_hash = EXCLUDED.api_key_hash,
        pc_name = EXCLUDED.pc_name,
        model_name = EXCLUDED.model_name,
        version = EXCLUDED.version,
        status = EXCLUDED.status,
        capabilities = EXCLUDED.capabilities,
        meta = EXCLUDED.meta,
        last_seen_at = now(),
        updated_at = now()
      RETURNING *
    `,
    [
      String(agent.agent_name || "").trim(),
      hashLlmAgentKey(apiKey),
      String(agent.pc_name || "").trim(),
      String(agent.model_name || "").trim(),
      String(agent.version || "").trim(),
      String(agent.status || "online").trim() || "online",
      JSON.stringify(Array.isArray(agent.capabilities) ? agent.capabilities : []),
      JSON.stringify(agent.meta && typeof agent.meta === "object" ? agent.meta : {}),
    ],
  );
  return mapLlmAgentRow(result.rows?.[0] || {});
}

export async function claimNextLlmJob(agentName, apiKey, options = {}) {
  if (!pool || !String(agentName || "").trim() || !String(apiKey || "").trim()) {
    return null;
  }
  const expectedHash = hashLlmAgentKey(apiKey);
  const client = await pool.connect();
  try {
    await client.query("BEGIN");
    const agentResult = await client.query(
      `
        SELECT *
        FROM llm_agents
        WHERE agent_name = $1 AND api_key_hash = $2
        FOR UPDATE
      `,
      [String(agentName || "").trim(), expectedHash],
    );
    const agentRow = agentResult.rows?.[0];
    if (!agentRow) {
      await client.query("ROLLBACK");
      return null;
    }

    const jobResult = await client.query(
      `
        SELECT *
        FROM llm_jobs
        WHERE status = 'pending'
          AND attempt_count < max_attempts
        ORDER BY priority DESC, created_at ASC
        LIMIT 1
        FOR UPDATE SKIP LOCKED
      `,
    );
    const jobRow = jobResult.rows?.[0];
    if (!jobRow) {
      await client.query(
        `
          UPDATE llm_agents
          SET status = $2, last_seen_at = now(), updated_at = now()
          WHERE agent_name = $1
        `,
        [String(agentName || "").trim(), String(options.agent_status || "idle").trim() || "idle"],
      );
      await client.query("COMMIT");
      return { agent: mapLlmAgentRow(agentRow), job: null };
    }

    const updatedJobResult = await client.query(
      `
        UPDATE llm_jobs
        SET
          status = 'claimed',
          agent_name = $2,
          claimed_at = now(),
          updated_at = now(),
          attempt_count = attempt_count + 1,
          error_text = ''
        WHERE id = $1
        RETURNING *
      `,
      [String(jobRow.id || "").trim(), String(agentName || "").trim()],
    );
    const updatedAgentResult = await client.query(
      `
        UPDATE llm_agents
        SET status = 'busy', last_seen_at = now(), last_job_claimed_at = now(), updated_at = now()
        WHERE agent_name = $1
        RETURNING *
      `,
      [String(agentName || "").trim()],
    );
    await client.query("COMMIT");
    return {
      agent: mapLlmAgentRow(updatedAgentResult.rows?.[0] || agentRow),
      job: mapLlmJobRow(updatedJobResult.rows?.[0] || {}),
    };
  } catch (error) {
    await client.query("ROLLBACK");
    throw error;
  } finally {
    client.release();
  }
}

export async function createLlmJob(job) {
  if (!pool) {
    throw new Error("Database is not initialized");
  }
  const jobId = String(job?.id || crypto.randomUUID()).trim();
  const result = await pool.query(
    `
      INSERT INTO llm_jobs (
        id, job_type, status, created_by, shipment_id, collection_id, document_kind,
        priority, attempt_count, max_attempts, agent_name, payload_json, result_json,
        error_text, created_at, updated_at
      )
      VALUES (
        $1, $2, 'pending', $3, $4, $5, $6,
        $7, 0, $8, '', $9::jsonb, '{}'::jsonb,
        '', now(), now()
      )
      RETURNING *
    `,
    [
      jobId,
      String(job?.job_type || "").trim(),
      String(job?.created_by || "").trim(),
      String(job?.shipment_id || "").trim(),
      String(job?.collection_id || "").trim(),
      String(job?.document_kind || "").trim(),
      Number(job?.priority || 0),
      Math.max(1, Number(job?.max_attempts || 3)),
      JSON.stringify(job?.payload_json && typeof job.payload_json === "object" ? job.payload_json : {}),
    ],
  );
  return mapLlmJobRow(result.rows?.[0] || {});
}

export async function completeLlmJob(jobId, agentName, resultJson) {
  if (!pool) {
    throw new Error("Database is not initialized");
  }
  const result = await pool.query(
    `
      UPDATE llm_jobs
      SET
        status = 'done',
        result_json = $3::jsonb,
        finished_at = now(),
        updated_at = now(),
        error_text = ''
      WHERE id = $1
        AND agent_name = $2
      RETURNING *
    `,
    [
      String(jobId || "").trim(),
      String(agentName || "").trim(),
      JSON.stringify(resultJson && typeof resultJson === "object" ? resultJson : {}),
    ],
  );
  return result.rows?.[0] ? mapLlmJobRow(result.rows[0]) : null;
}

export async function failLlmJob(jobId, agentName, errorText, allowRetry = false) {
  if (!pool) {
    throw new Error("Database is not initialized");
  }
  const result = await pool.query(
    `
      UPDATE llm_jobs
      SET
        status = CASE
          WHEN $4 = true AND attempt_count < max_attempts THEN 'pending'
          ELSE 'failed'
        END,
        agent_name = CASE
          WHEN $4 = true AND attempt_count < max_attempts THEN ''
          ELSE agent_name
        END,
        claimed_at = CASE
          WHEN $4 = true AND attempt_count < max_attempts THEN null
          ELSE claimed_at
        END,
        finished_at = CASE
          WHEN $4 = true AND attempt_count < max_attempts THEN null
          ELSE now()
        END,
        updated_at = now(),
        error_text = $3
      WHERE id = $1
        AND agent_name = $2
      RETURNING *
    `,
    [
      String(jobId || "").trim(),
      String(agentName || "").trim(),
      String(errorText || "").trim(),
      allowRetry === true,
    ],
  );
  return result.rows?.[0] ? mapLlmJobRow(result.rows[0]) : null;
}

export async function getLlmQueueSnapshot() {
  if (!pool) {
    return {
      agents: [],
      jobs: [],
      summary: {
        total_jobs: 0,
        pending_jobs: 0,
        claimed_jobs: 0,
        failed_jobs: 0,
        done_jobs: 0,
        online_agents: 0,
      },
    };
  }
  const [agentsResult, jobsResult, summaryResult] = await Promise.all([
    pool.query("SELECT * FROM llm_agents ORDER BY agent_name ASC"),
    pool.query("SELECT * FROM llm_jobs ORDER BY created_at DESC LIMIT 50"),
    pool.query(`
      SELECT
        COUNT(*)::int AS total_jobs,
        COUNT(*) FILTER (WHERE status = 'pending')::int AS pending_jobs,
        COUNT(*) FILTER (WHERE status = 'claimed')::int AS claimed_jobs,
        COUNT(*) FILTER (WHERE status = 'failed')::int AS failed_jobs,
        COUNT(*) FILTER (WHERE status = 'done')::int AS done_jobs,
        (SELECT COUNT(*)::int FROM llm_agents WHERE status IN ('online', 'idle', 'busy')) AS online_agents
      FROM llm_jobs
    `),
  ]);
  return {
    agents: (agentsResult.rows || []).map(mapLlmAgentRow),
    jobs: (jobsResult.rows || []).map(mapLlmJobRow),
    summary: {
      total_jobs: Number(summaryResult.rows?.[0]?.total_jobs || 0),
      pending_jobs: Number(summaryResult.rows?.[0]?.pending_jobs || 0),
      claimed_jobs: Number(summaryResult.rows?.[0]?.claimed_jobs || 0),
      failed_jobs: Number(summaryResult.rows?.[0]?.failed_jobs || 0),
      done_jobs: Number(summaryResult.rows?.[0]?.done_jobs || 0),
      online_agents: Number(summaryResult.rows?.[0]?.online_agents || 0),
    },
  };
}

const databaseMigrations = [
  `
    CREATE TABLE IF NOT EXISTS fust_actions (
      id text PRIMARY KEY,
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
      confirmed_at timestamptz,
      confirmed_by text,
      import_source jsonb NOT NULL DEFAULT '{}'::jsonb,
      confirmation_reminder jsonb NOT NULL DEFAULT '{}'::jsonb,
      updated_at timestamptz NOT NULL DEFAULT now()
    )
  `,
  `
    ALTER TABLE IF EXISTS fust_actions
    ADD COLUMN IF NOT EXISTS confirmed_at timestamptz
  `,
  `
    ALTER TABLE IF EXISTS fust_actions
    ADD COLUMN IF NOT EXISTS confirmed_by text
  `,
  `
    ALTER TABLE IF EXISTS fust_actions
    ADD COLUMN IF NOT EXISTS import_source jsonb NOT NULL DEFAULT '{}'::jsonb
  `,
  `
    ALTER TABLE IF EXISTS fust_actions
    ADD COLUMN IF NOT EXISTS confirmation_reminder jsonb NOT NULL DEFAULT '{}'::jsonb
  `,
  `
    CREATE TABLE IF NOT EXISTS fust_action_documents (
      id bigserial PRIMARY KEY,
      action_id text NOT NULL REFERENCES fust_actions(id) ON DELETE CASCADE,
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
  `
    ALTER TABLE IF EXISTS fust_action_documents
    DROP CONSTRAINT IF EXISTS fust_action_documents_action_id_fkey
  `,
  `
    ALTER TABLE IF EXISTS fust_action_documents
    ALTER COLUMN action_id TYPE text USING action_id::text
  `,
  `
    ALTER TABLE IF EXISTS fust_actions
    ALTER COLUMN id TYPE text USING id::text
  `,
  `
    DO $$
    BEGIN
      IF NOT EXISTS (
        SELECT 1
        FROM pg_constraint
        WHERE conname = 'fust_action_documents_action_id_fkey'
      ) THEN
        ALTER TABLE fust_action_documents
        ADD CONSTRAINT fust_action_documents_action_id_fkey
        FOREIGN KEY (action_id) REFERENCES fust_actions(id) ON DELETE CASCADE;
      END IF;
    END $$;
  `,
  `
    CREATE TABLE IF NOT EXISTS llm_agents (
      agent_name text PRIMARY KEY,
      api_key_hash text NOT NULL,
      pc_name text,
      model_name text,
      version text,
      status text NOT NULL DEFAULT 'offline',
      capabilities jsonb NOT NULL DEFAULT '[]'::jsonb,
      meta jsonb NOT NULL DEFAULT '{}'::jsonb,
      last_seen_at timestamptz,
      last_job_claimed_at timestamptz,
      created_at timestamptz NOT NULL DEFAULT now(),
      updated_at timestamptz NOT NULL DEFAULT now()
    )
  `,
  `
    CREATE TABLE IF NOT EXISTS llm_jobs (
      id text PRIMARY KEY,
      job_type text NOT NULL,
      status text NOT NULL DEFAULT 'pending',
      created_by text,
      shipment_id text,
      collection_id text,
      document_kind text,
      priority integer NOT NULL DEFAULT 0,
      attempt_count integer NOT NULL DEFAULT 0,
      max_attempts integer NOT NULL DEFAULT 3,
      agent_name text NOT NULL DEFAULT '',
      payload_json jsonb NOT NULL DEFAULT '{}'::jsonb,
      result_json jsonb NOT NULL DEFAULT '{}'::jsonb,
      error_text text NOT NULL DEFAULT '',
      created_at timestamptz NOT NULL DEFAULT now(),
      claimed_at timestamptz,
      finished_at timestamptz,
      updated_at timestamptz NOT NULL DEFAULT now()
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
