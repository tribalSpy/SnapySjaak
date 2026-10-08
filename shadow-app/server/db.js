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
  const created = new Pool({
    connectionString,
    ssl: sslEnabled ? { rejectUnauthorized: false } : false,
    // Detect dropped connections sooner instead of finding out on the next query.
    keepAlive: true,
  });
  // A connection the database drops (Postgres restarting / recovering,
  // network blip) emits an 'error' event; with no listener, Node treats it
  // as uncaught and the WHOLE server crashes ("Connection terminated
  // unexpectedly" -> "throw er; // Unhandled 'error' event", confirmed on
  // Render). Logged instead: the pool discards that connection and opens a
  // new one on the next query. Idle connections report on the pool,
  // checked-out ones on the client itself -- both need a listener.
  created.on("error", (error) => {
    console.error("Database connection lost (idle):", error?.message || error);
  });
  created.on("connect", (client) => {
    client.on("error", (error) => {
      console.error("Database connection lost (in use):", error?.message || error);
    });
  });
  return created;
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
        deleted, deleted_at, deleted_by, created_by, created_at, confirmed_at, confirmed_by,
        import_source, confirmation_reminder, sheet_sync, email_sync, db_sync, updated_at
      )
      VALUES (
        $1, $2, $3, $4, $5, $6, $7, $8, $9,
        $10, $11, $12, $13, $14, $15, $16, $17, $18,
        $19, $20, $21, $22, COALESCE($23, now()), $24, $25,
        $26::jsonb, $27::jsonb, $28::jsonb, $29::jsonb, $30::jsonb, now()
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
        deleted_at = EXCLUDED.deleted_at,
        deleted_by = EXCLUDED.deleted_by,
        created_by = EXCLUDED.created_by,
        created_at = COALESCE(fust_actions.created_at, EXCLUDED.created_at),
        confirmed_at = EXCLUDED.confirmed_at,
        confirmed_by = EXCLUDED.confirmed_by,
        import_source = EXCLUDED.import_source,
        confirmation_reminder = EXCLUDED.confirmation_reminder,
        sheet_sync = EXCLUDED.sheet_sync,
        email_sync = EXCLUDED.email_sync,
        db_sync = EXCLUDED.db_sync,
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
      action.deleted_at || null,
      action.deleted_by || "",
      action.created_by || "",
      action.created_at || null,
      action.confirmed_at || null,
      action.confirmed_by || "",
      JSON.stringify(action.import_source || {}),
      JSON.stringify(action.confirmation_reminder || {}),
      JSON.stringify(action.sheet_sync || {}),
      JSON.stringify(action.email_sync || {}),
      JSON.stringify(action.db_sync || {}),
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

function numberOrNull(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : null;
}

function chunkArray(items, size) {
  const chunks = [];
  for (let index = 0; index < items.length; index += size) {
    chunks.push(items.slice(index, index + size));
  }
  return chunks;
}

// Postgres refuses "ON CONFLICT DO UPDATE command cannot affect row a
// second time" the moment two rows in the SAME multi-row INSERT share a
// conflict target -- confirmed real: a large enough ERP/dispatch export can
// legitimately repeat the same lot number (e.g. a lot still open across
// several of the export's covered days). Keeping the LAST occurrence for a
// given key (consistent with every other re-upload here: newer data wins)
// guarantees no batch can ever trigger it, regardless of why the source
// file happened to repeat a key.
function dedupeByKeyKeepingLast(items, keyFn) {
  const byKey = new Map();
  for (const item of items) {
    byKey.set(keyFn(item), item);
  }
  return [...byKey.values()];
}

// The ERP export alone runs ~1600+ rows a day -- one INSERT per row (the
// fust_reference_actions pattern above, fine at its tens-of-codes daily
// volume) would mean thousands of sequential round-trips inside the HTTP
// request the user is waiting on. These batch every chunk into a single
// multi-row INSERT ... ON CONFLICT instead.
const INKOOP_UPSERT_CHUNK_SIZE = 500;

export async function saveInkoopErpLines(rows) {
  if (!pool || !Array.isArray(rows) || !rows.length) {
    return;
  }
  const validRows = dedupeByKeyKeepingLast(
    rows.filter((row) => numberOrNull(row?.lot) !== null),
    (row) => numberOrNull(row.lot),
  );
  for (const chunk of chunkArray(validRows, INKOOP_UPSERT_CHUNK_SIZE)) {
    const values = [];
    const placeholders = chunk.map((row) => {
      const base = values.length;
      values.push(
        numberOrNull(row.lot),
        row.date || null,
        row.pav || "",
        row.suppl || "",
        row.transp || "",
        row.avc || "",
        row.description || "",
        numberOrNull(row.pieces),
        numberOrNull(row.price),
        numberOrNull(row.t_price),
        row.deb_no || "",
        row.inv_no || "",
        JSON.stringify(row),
      );
      return `($${base + 1}, $${base + 2}::date, $${base + 3}, $${base + 4}, $${base + 5}, $${base + 6}, $${base + 7}, $${base + 8}, $${base + 9}, $${base + 10}, $${base + 11}, $${base + 12}, $${base + 13}::jsonb)`;
    });
    await pool.query(
      `
        INSERT INTO inkoop_erp_lines (
          lot, erp_date, pav, suppl, transp, avc, description, pieces, price, t_price, deb_no, inv_no, raw
        )
        VALUES ${placeholders.join(", ")}
        ON CONFLICT (lot) DO UPDATE SET
          erp_date = EXCLUDED.erp_date,
          pav = EXCLUDED.pav,
          suppl = EXCLUDED.suppl,
          transp = EXCLUDED.transp,
          avc = EXCLUDED.avc,
          description = EXCLUDED.description,
          pieces = EXCLUDED.pieces,
          price = EXCLUDED.price,
          t_price = EXCLUDED.t_price,
          deb_no = EXCLUDED.deb_no,
          inv_no = EXCLUDED.inv_no,
          raw = EXCLUDED.raw,
          updated_at = now()
      `,
      values,
    );
  }
}

export async function saveInkoopInvoiceLines(lines) {
  if (!pool || !Array.isArray(lines) || !lines.length) {
    return;
  }
  const validLines = lines.filter((line) => line?.id);
  for (const chunk of chunkArray(validLines, INKOOP_UPSERT_CHUNK_SIZE)) {
    const values = [];
    const placeholders = chunk.map((line) => {
      const base = values.length;
      values.push(
        line.id,
        line.invoice_number || "",
        line.invoice_type || "",
        line.reference_bt || "",
        line.supplier_gln || "",
        line.supplier_fh_number || "",
        line.supplier_name || "",
        line.description || "",
        numberOrNull(line.quantity),
        numberOrNull(line.unit_price),
        numberOrNull(line.total),
        JSON.stringify(line),
        line.source_file_name || "",
        line.invoice_date || null,
        line.company_number || "",
        line.company_name || "",
      );
      return `($${base + 1}, $${base + 2}, $${base + 3}, $${base + 4}, $${base + 5}, $${base + 6}, $${base + 7}, $${base + 8}, $${base + 9}, $${base + 10}, $${base + 11}, $${base + 12}::jsonb, $${base + 13}, $${base + 14}::date, $${base + 15}, $${base + 16})`;
    });
    await pool.query(
      `
        INSERT INTO inkoop_invoice_lines (
          id, invoice_number, invoice_type, reference_bt, supplier_gln, supplier_fh_number, supplier_name,
          description, quantity, unit_price, total, raw, source_file_name, invoice_date, company_number, company_name
        )
        VALUES ${placeholders.join(", ")}
        ON CONFLICT (id) DO UPDATE SET
          invoice_number = EXCLUDED.invoice_number,
          invoice_type = EXCLUDED.invoice_type,
          reference_bt = EXCLUDED.reference_bt,
          supplier_gln = EXCLUDED.supplier_gln,
          supplier_fh_number = EXCLUDED.supplier_fh_number,
          supplier_name = EXCLUDED.supplier_name,
          description = EXCLUDED.description,
          quantity = EXCLUDED.quantity,
          unit_price = EXCLUDED.unit_price,
          total = EXCLUDED.total,
          raw = EXCLUDED.raw,
          source_file_name = EXCLUDED.source_file_name,
          invoice_date = EXCLUDED.invoice_date,
          company_number = EXCLUDED.company_number,
          company_name = EXCLUDED.company_name,
          updated_at = now()
      `,
      values,
    );
  }
}

export async function saveInkoopDispatchLots(lots) {
  if (!pool || !Array.isArray(lots) || !lots.length) {
    return;
  }
  // parse_dispatch_dump already merges a recurring lot's dispatches into one
  // record -- this is only a backstop, so keeping the last occurrence here
  // (rather than merging again) is fine.
  const validLots = dedupeByKeyKeepingLast(
    lots.filter((lot) => numberOrNull(lot?.lot) !== null),
    (lot) => numberOrNull(lot.lot),
  );
  for (const chunk of chunkArray(validLots, INKOOP_UPSERT_CHUNK_SIZE)) {
    const values = [];
    const placeholders = chunk.map((lot) => {
      const base = values.length;
      values.push(
        numberOrNull(lot.lot),
        lot.date || null,
        lot.pav || "",
        lot.suppl || "",
        lot.description || "",
        numberOrNull(lot.pieces),
        numberOrNull(lot.price),
        numberOrNull(lot.t_price),
        JSON.stringify(Array.isArray(lot.dispatches) ? lot.dispatches : []),
      );
      return `($${base + 1}, $${base + 2}::date, $${base + 3}, $${base + 4}, $${base + 5}, $${base + 6}, $${base + 7}, $${base + 8}, $${base + 9}::jsonb)`;
    });
    await pool.query(
      `
        INSERT INTO inkoop_dispatch_lots (
          lot, erp_date, pav, suppl, description, pieces, price, t_price, dispatches
        )
        VALUES ${placeholders.join(", ")}
        ON CONFLICT (lot) DO UPDATE SET
          erp_date = EXCLUDED.erp_date,
          pav = EXCLUDED.pav,
          suppl = EXCLUDED.suppl,
          description = EXCLUDED.description,
          pieces = EXCLUDED.pieces,
          price = EXCLUDED.price,
          t_price = EXCLUDED.t_price,
          dispatches = EXCLUDED.dispatches,
          updated_at = now()
      `,
      values,
    );
  }
}

export async function getInkoopDispatchLots({ from, to } = {}) {
  if (!pool) {
    return [];
  }
  const result = await pool.query(
    `
      SELECT lot, erp_date, pav, suppl, description, pieces, price, t_price, dispatches
      FROM inkoop_dispatch_lots
      WHERE erp_date IS NULL OR (erp_date >= $1 AND erp_date <= $2)
      ORDER BY lot
    `,
    [from, to],
  );
  return result.rows;
}

export async function getInkoopDispatchLotsByLots(lotNumbers) {
  if (!pool || !Array.isArray(lotNumbers) || !lotNumbers.length) {
    return [];
  }
  const validLots = [...new Set(lotNumbers.map((lot) => numberOrNull(lot)).filter((lot) => lot !== null))];
  if (!validLots.length) {
    return [];
  }
  const result = await pool.query(
    `
      SELECT lot, erp_date, pav, suppl, description, pieces, price, t_price, dispatches
      FROM inkoop_dispatch_lots
      WHERE lot = ANY($1::bigint[])
    `,
    [validLots],
  );
  return result.rows;
}

export async function upsertShelfCountRow(row) {
  if (!pool || !row?.customer_reference || !row?.nightly_run_date) {
    return null;
  }
  const result = await pool.query(
    `
      INSERT INTO shelf_counts (
        customer_reference, nightly_run_date, drive_folder_name, trolley_count, photo_count, shelf_count,
        level_count, confidence, expected_average, deviation, extension_count, extension_expected,
        extension_deviation, status, model_version, job_id, error_text, processed_at, per_trolley, trained_model, expected_trolleys, combined_into
      )
      VALUES ($1, $2::date, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18::timestamptz, $19::jsonb, $20::jsonb, $21, $22)
      ON CONFLICT (customer_reference, nightly_run_date) DO UPDATE SET
        drive_folder_name = COALESCE(NULLIF(EXCLUDED.drive_folder_name, ''), shelf_counts.drive_folder_name),
        trolley_count = COALESCE(NULLIF(EXCLUDED.trolley_count, 0), shelf_counts.trolley_count),
        photo_count = COALESCE(NULLIF(EXCLUDED.photo_count, 0), shelf_counts.photo_count),
        shelf_count = COALESCE(EXCLUDED.shelf_count, shelf_counts.shelf_count),
        level_count = COALESCE(EXCLUDED.level_count, shelf_counts.level_count),
        confidence = COALESCE(EXCLUDED.confidence, shelf_counts.confidence),
        expected_average = COALESCE(EXCLUDED.expected_average, shelf_counts.expected_average),
        deviation = COALESCE(EXCLUDED.deviation, shelf_counts.deviation),
        extension_count = COALESCE(EXCLUDED.extension_count, shelf_counts.extension_count),
        extension_expected = COALESCE(EXCLUDED.extension_expected, shelf_counts.extension_expected),
        extension_deviation = COALESCE(EXCLUDED.extension_deviation, shelf_counts.extension_deviation),
        status = EXCLUDED.status,
        model_version = COALESCE(NULLIF(EXCLUDED.model_version, ''), shelf_counts.model_version),
        job_id = COALESCE(NULLIF(EXCLUDED.job_id, ''), shelf_counts.job_id),
        error_text = EXCLUDED.error_text,
        processed_at = COALESCE(EXCLUDED.processed_at, shelf_counts.processed_at),
        per_trolley = COALESCE(EXCLUDED.per_trolley, shelf_counts.per_trolley),
        trained_model = COALESCE(EXCLUDED.trained_model, shelf_counts.trained_model),
        expected_trolleys = COALESCE(EXCLUDED.expected_trolleys, shelf_counts.expected_trolleys),
        combined_into = COALESCE(NULLIF(EXCLUDED.combined_into, ''), shelf_counts.combined_into),
        updated_at = now()
      RETURNING *
    `,
    [
      row.customer_reference,
      row.nightly_run_date,
      row.drive_folder_name || "",
      numberOrNull(row.trolley_count) || 0,
      numberOrNull(row.photo_count) || 0,
      numberOrNull(row.shelf_count),
      numberOrNull(row.level_count),
      numberOrNull(row.confidence),
      numberOrNull(row.expected_average),
      numberOrNull(row.deviation),
      numberOrNull(row.extension_count),
      numberOrNull(row.extension_expected),
      numberOrNull(row.extension_deviation),
      row.status || "pending",
      row.model_version || "",
      row.job_id || "",
      row.error_text || "",
      row.processed_at || null,
      Array.isArray(row.per_trolley) ? JSON.stringify(row.per_trolley) : null,
      row.trained_model && typeof row.trained_model === "object" ? JSON.stringify(row.trained_model) : null,
      numberOrNull(row.expected_trolleys),
      row.combined_into || "",
    ],
  );
  return result.rows?.[0] || null;
}

// Writes the Fust Planning expectations exactly as given -- a value Fust
// doesn't have (null) is cleared, not kept. upsertShelfCountRow COALESCEs,
// which left an old value standing (confirmed: "Fust shelves" kept showing
// the trolley count from before the DC/DCS fix when DCS was empty).
export async function setShelfCountExpectations(customerReference, date, values) {
  if (!pool || !customerReference || !date) {
    return;
  }
  await pool.query(
    `
      UPDATE shelf_counts
      SET expected_average = $3, deviation = $4, extension_expected = $5, extension_deviation = $6,
          expected_trolleys = $7, updated_at = now()
      WHERE customer_reference = $1 AND nightly_run_date = $2::date
    `,
    [
      customerReference,
      date,
      numberOrNull(values.expected_average),
      numberOrNull(values.deviation),
      numberOrNull(values.extension_expected),
      numberOrNull(values.extension_deviation),
      numberOrNull(values.expected_trolleys),
    ],
  );
}

export async function getShelfCountsForDate(date) {
  if (!pool || !date) {
    return [];
  }
  const result = await pool.query(
    "SELECT * FROM shelf_counts WHERE nightly_run_date = $1::date ORDER BY customer_reference",
    [date],
  );
  return result.rows;
}

// Feeds the training pipeline's dataset-collection script (shelf-training/
// data/collect.py) -- a standalone tool run on the GPU PC, kept as an HTTP
// client of shadow-app rather than a second thing with direct Postgres
// credentials, the same way the pollers never touch the database directly
// either. statusFilter narrows to specific statuses (e.g. ["needs_review"]
// for the "hard cases" priority tier); omitted, it returns every status.
export async function getShelfCountsInRange({ from, to, statusFilter } = {}) {
  if (!pool || !from || !to) {
    return [];
  }
  const statuses = Array.isArray(statusFilter) && statusFilter.length ? statusFilter : null;
  const result = await pool.query(
    `
      SELECT * FROM shelf_counts
      WHERE nightly_run_date >= $1::date AND nightly_run_date <= $2::date
        AND ($3::text[] IS NULL OR status = ANY($3::text[]))
      ORDER BY nightly_run_date, customer_reference
    `,
    [from, to, statuses],
  );
  return result.rows;
}

export async function upsertShelfCountNightlyRun(run) {
  if (!pool || !run?.run_date) {
    return null;
  }
  const result = await pool.query(
    `
      INSERT INTO shelf_count_nightly_runs (run_date, status, totals, issues, started_at, completed_at)
      VALUES ($1::date, $2, $3::jsonb, $4::jsonb, COALESCE($5::timestamptz, now()), $6::timestamptz)
      ON CONFLICT (run_date) DO UPDATE SET
        status = EXCLUDED.status,
        -- Merged, not replaced: the trigger only writes its own keys
        -- (checked/missing_photos/unmatched_folder) while finished jobs
        -- increment ok/needs_review/failed concurrently -- replacing used to
        -- wipe those. reset_totals (a fresh/explicit re-run) starts over.
        totals = CASE WHEN $7 THEN EXCLUDED.totals
                      ELSE COALESCE(shelf_count_nightly_runs.totals, '{}'::jsonb) || EXCLUDED.totals END,
        issues = EXCLUDED.issues,
        started_at = CASE WHEN $7 THEN EXCLUDED.started_at ELSE shelf_count_nightly_runs.started_at END,
        completed_at = COALESCE(EXCLUDED.completed_at, shelf_count_nightly_runs.completed_at),
        updated_at = now()
      RETURNING *
    `,
    [
      run.run_date,
      run.status || "running",
      JSON.stringify(run.totals && typeof run.totals === "object" ? run.totals : {}),
      JSON.stringify(Array.isArray(run.issues) ? run.issues : []),
      run.started_at || null,
      run.completed_at || null,
      run.reset_totals === true,
    ],
  );
  return result.rows?.[0] || null;
}

// Runs still marked "running" -- the watchdog resumes any whose updated_at
// (bumped after every reference the trigger handles) has gone quiet, i.e.
// the server crashed or restarted mid-run.
export async function getRunningShelfCountNightlyRuns() {
  if (!pool) {
    return [];
  }
  const result = await pool.query(
    "SELECT to_char(run_date, 'YYYY-MM-DD') AS run_date, status, started_at, updated_at FROM shelf_count_nightly_runs WHERE status = 'running' ORDER BY run_date DESC",
  );
  return result.rows;
}

export async function getShelfCountNightlyRun(date) {
  if (!pool || !date) {
    return null;
  }
  const result = await pool.query("SELECT * FROM shelf_count_nightly_runs WHERE run_date = $1::date", [date]);
  return result.rows?.[0] || null;
}

// Atomic increment via jsonb_set/a computed expression, not a JS read-
// modify-write -- multiple shelf_count jobs for the same night can complete
// concurrently (once more than one poller exists), and a read-modify-write
// here would silently drop whichever completion's write lost the race.
export async function incrementShelfCountNightlyRunTotal(runDate, statusKey) {
  if (!pool || !runDate || !statusKey) {
    return null;
  }
  const result = await pool.query(
    `
      UPDATE shelf_count_nightly_runs
      SET totals = jsonb_set(
            COALESCE(totals, '{}'::jsonb),
            ARRAY[$2],
            to_jsonb(COALESCE((totals->>$2)::int, 0) + 1)
          ),
          updated_at = now()
      WHERE run_date = $1::date
      RETURNING *
    `,
    [runDate, statusKey],
  );
  return result.rows?.[0] || null;
}

export async function getInkoopErpLines({ from, to } = {}) {
  if (!pool) {
    return [];
  }
  const result = await pool.query(
    `
      SELECT raw FROM inkoop_erp_lines
      WHERE erp_date IS NULL OR (erp_date >= $1 AND erp_date <= $2)
      ORDER BY lot
    `,
    [from, to],
  );
  return result.rows.map((row) => row.raw);
}

export async function getInkoopInvoiceLines({ from, to } = {}) {
  if (!pool) {
    return [];
  }
  const result = await pool.query(
    `
      SELECT raw FROM inkoop_invoice_lines
      WHERE invoice_date IS NULL OR (invoice_date >= $1 AND invoice_date <= $2)
      ORDER BY invoice_number
    `,
    [from, to],
  );
  return result.rows.map((row) => row.raw);
}

// "Start fresh": every uploaded Inkoop Controle record -- ERP ledger,
// invoice ledger, dispatch dump, Follow-up statuses/notes. Master data and
// manual links live in inkoop-state.json and are deliberately not touched.
export async function clearInkoopUploads() {
  if (!pool) {
    return { cleared: false };
  }
  const counts = {};
  for (const table of ["inkoop_erp_lines", "inkoop_invoice_lines", "inkoop_invoice_headers", "inkoop_dispatch_lots", "inkoop_issue_status"]) {
    const result = await pool.query(`SELECT COUNT(*)::int AS count FROM ${table}`);
    counts[table] = result.rows[0]?.count || 0;
  }
  await pool.query("TRUNCATE inkoop_erp_lines, inkoop_invoice_lines, inkoop_invoice_headers, inkoop_dispatch_lots, inkoop_issue_status");
  return { cleared: true, counts };
}

export async function saveInkoopInvoiceHeaders(headers) {
  if (!pool || !Array.isArray(headers) || !headers.length) {
    return;
  }
  for (const header of headers.filter((item) => item?.invoice_number)) {
    await pool.query(
      `
        INSERT INTO inkoop_invoice_headers (
          invoice_number, invoice_type, company_number, company_name, invoice_date, grand_total, currency,
          vat_subtotals, klok_location, klok_locations, source_file_name
        )
        VALUES ($1, $2, $3, $4, $5::date, $6, $7, $8::jsonb, $9, $10::jsonb, $11)
        ON CONFLICT (invoice_number) DO UPDATE SET
          invoice_type = EXCLUDED.invoice_type,
          company_number = EXCLUDED.company_number,
          company_name = EXCLUDED.company_name,
          invoice_date = EXCLUDED.invoice_date,
          grand_total = EXCLUDED.grand_total,
          currency = EXCLUDED.currency,
          vat_subtotals = EXCLUDED.vat_subtotals,
          klok_location = EXCLUDED.klok_location,
          klok_locations = EXCLUDED.klok_locations,
          source_file_name = EXCLUDED.source_file_name,
          updated_at = now()
      `,
      [
        header.invoice_number,
        header.invoice_type || "",
        header.company_number || "",
        header.company_name || "",
        header.invoice_date || null,
        numberOrNull(header.grand_total),
        header.currency || "EUR",
        JSON.stringify(Array.isArray(header.vat_subtotals) ? header.vat_subtotals : []),
        header.klok_location || "",
        JSON.stringify(Array.isArray(header.klok_locations) ? header.klok_locations : []),
        header.source_file_name || "",
      ],
    );
  }
}

export async function getInkoopInvoiceHeaders({ from, to } = {}) {
  if (!pool || !from || !to) {
    return [];
  }
  const result = await pool.query(
    `
      SELECT invoice_number, invoice_type, company_number, company_name,
        to_char(invoice_date, 'YYYY-MM-DD') AS invoice_date, grand_total::float AS grand_total, currency,
        vat_subtotals, klok_location, klok_locations, source_file_name
      FROM inkoop_invoice_headers
      WHERE invoice_date >= $1::date AND invoice_date <= $2::date
      ORDER BY company_number, invoice_number
    `,
    [from, to],
  );
  return result.rows;
}

// The invoices with lines on `date` (the line/auction date the calendar and
// the compare use -- not the invoice's own issue date), with their header
// when one was stored (invoices uploaded before the King export have none).
export async function getInkoopInvoicesByLineDate(date) {
  if (!pool || !date) {
    return [];
  }
  const result = await pool.query(
    `
      SELECT l.invoice_number, MIN(l.invoice_type) AS invoice_type, MIN(l.company_number) AS company_number,
        MIN(l.company_name) AS company_name,
        h.invoice_number IS NOT NULL AS has_header,
        to_char(h.invoice_date, 'YYYY-MM-DD') AS invoice_date, h.grand_total::float AS grand_total, h.currency,
        h.vat_subtotals, h.klok_location, h.klok_locations, h.source_file_name, h.company_name AS header_company_name
      FROM inkoop_invoice_lines l
      LEFT JOIN inkoop_invoice_headers h ON h.invoice_number = l.invoice_number
      WHERE l.invoice_date = $1::date AND l.invoice_number <> ''
      GROUP BY l.invoice_number, h.invoice_number
      ORDER BY MIN(l.company_number), l.invoice_number
    `,
    [date],
  );
  return result.rows;
}

// Every stored line of the given invoices (the raw jsonb carries the King
// fields: prd_id, vat_category, trigger_code, ...).
export async function getInkoopInvoiceLinesByInvoiceNumbers(invoiceNumbers) {
  if (!pool || !Array.isArray(invoiceNumbers) || !invoiceNumbers.length) {
    return [];
  }
  const result = await pool.query(
    "SELECT raw FROM inkoop_invoice_lines WHERE invoice_number = ANY($1::text[]) ORDER BY invoice_number, id",
    [invoiceNumbers],
  );
  return result.rows.map((row) => row.raw);
}

export async function getKingExports(invoiceNumbers = null) {
  if (!pool) {
    return [];
  }
  const result = Array.isArray(invoiceNumbers)
    ? await pool.query(
      `SELECT invoice_number, batch_id, stuknummer, extern_id, status, job_id, error_text, exported_by,
         exported_at, delivered_at FROM king_exports WHERE invoice_number = ANY($1::text[])`,
      [invoiceNumbers],
    )
    : await pool.query(
      `SELECT invoice_number, batch_id, stuknummer, extern_id, status, job_id, error_text, exported_by,
         exported_at, delivered_at FROM king_exports ORDER BY exported_at DESC LIMIT 500`,
    );
  return result.rows;
}

export async function saveKingExports(rows) {
  if (!pool || !Array.isArray(rows) || !rows.length) {
    return;
  }
  for (const row of rows) {
    await pool.query(
      `
        INSERT INTO king_exports (invoice_number, batch_id, stuknummer, extern_id, status, job_id, error_text, exported_by, exported_at, delivered_at, journal_json)
        VALUES ($1, $2, $3, $4, $5, $6, $7, $8, now(), NULL, $9::jsonb)
        ON CONFLICT (invoice_number) DO UPDATE SET
          batch_id = EXCLUDED.batch_id,
          stuknummer = EXCLUDED.stuknummer,
          extern_id = EXCLUDED.extern_id,
          status = EXCLUDED.status,
          job_id = EXCLUDED.job_id,
          error_text = EXCLUDED.error_text,
          exported_by = EXCLUDED.exported_by,
          exported_at = now(),
          delivered_at = NULL,
          journal_json = EXCLUDED.journal_json
      `,
      [
        row.invoice_number,
        row.batch_id,
        Number.isFinite(Number(row.stuknummer)) ? Number(row.stuknummer) : null,
        row.extern_id || "",
        row.status || "queued",
        row.job_id || "",
        row.error_text || "",
        row.exported_by || "",
        JSON.stringify(row.journal_json || null),
      ],
    );
  }
}

export async function updateKingExportBatchStatus(batchId, status, errorText = "") {
  if (!pool || !batchId) {
    return;
  }
  await pool.query(
    `UPDATE king_exports SET status = $2, error_text = $3,
       delivered_at = CASE WHEN $2 = 'delivered' THEN now() ELSE delivered_at END
     WHERE batch_id = $1`,
    [batchId, status, errorText],
  );
}

// Just the columns the per-company invoice summary needs (not the full raw
// line), for every line in the date range.
export async function getInkoopInvoiceSummaryLines({ from, to } = {}) {
  if (!pool || !from || !to) {
    return [];
  }
  const result = await pool.query(
    `
      SELECT company_number, company_name, invoice_number, invoice_type, to_char(invoice_date, 'YYYY-MM-DD') AS invoice_date, description, quantity, total
      FROM inkoop_invoice_lines
      WHERE invoice_date >= $1::date AND invoice_date <= $2::date
    `,
    [from, to],
  );
  return result.rows;
}

export async function getInkoopCompanies() {
  if (!pool) {
    return [];
  }
  const result = await pool.query(
    `
      SELECT company_number, MAX(company_name) AS company_name
      FROM inkoop_invoice_lines
      WHERE company_number IS NOT NULL AND company_number <> ''
      GROUP BY company_number
      ORDER BY company_number
    `,
  );
  return result.rows;
}

export async function saveInkoopIssueStatus(issue) {
  if (!pool || !issue?.id) {
    return;
  }
  await pool.query(
    `
      INSERT INTO inkoop_issue_status (id, issue_type, status, assigned_to, note, updated_by, updated_at)
      VALUES ($1, $2, $3, $4, $5, $6, now())
      ON CONFLICT (id) DO UPDATE SET
        issue_type = EXCLUDED.issue_type,
        status = EXCLUDED.status,
        assigned_to = EXCLUDED.assigned_to,
        note = EXCLUDED.note,
        updated_by = EXCLUDED.updated_by,
        updated_at = now()
    `,
    [
      issue.id,
      issue.issue_type || "",
      issue.status || "open",
      issue.assigned_to || "",
      issue.note || "",
      issue.updated_by || "",
    ],
  );
}

export async function getInkoopIssueStatuses() {
  if (!pool) {
    return [];
  }
  const result = await pool.query("SELECT * FROM inkoop_issue_status");
  return result.rows;
}

// Reference-code-level rows from the Fust API import (svdvyver.fr) -- kept
// separate from fust_actions since a Code isn't a real customer for
// CMR/Fustbon/confirmation purposes. A null metric here means "no data
// yet" (the source only fills DC-Actual/DCS/etc. in once a day closes
// out), so every numeric/boolean column is nullable with no default.
export async function saveFustReferenceAction(row) {
  if (!pool || !row?.id) {
    return;
  }
  await pool.query(
    `
      INSERT INTO fust_reference_actions (
        id, action_date, country, code, group_with, carrier1_name, carrier2_name,
        matched_customer_name, matched_connect_name, matched_customer_code, matched_by,
        hours, flo, pla, acc, all_flag, boxes,
        dc_planning, dc_actual, dcs, dco, cctag, vk, pal, raw, updated_at
      )
      VALUES (
        $1, $2, $3, $4, $5, $6, $7,
        $8, $9, $10, $11,
        $12, $13, $14, $15, $16, $17,
        $18, $19, $20, $21, $22, $23, $24, $25::jsonb, now()
      )
      ON CONFLICT (id) DO UPDATE SET
        action_date = EXCLUDED.action_date,
        country = EXCLUDED.country,
        code = EXCLUDED.code,
        group_with = EXCLUDED.group_with,
        carrier1_name = EXCLUDED.carrier1_name,
        carrier2_name = EXCLUDED.carrier2_name,
        matched_customer_name = EXCLUDED.matched_customer_name,
        matched_connect_name = EXCLUDED.matched_connect_name,
        matched_customer_code = EXCLUDED.matched_customer_code,
        matched_by = EXCLUDED.matched_by,
        hours = EXCLUDED.hours,
        flo = EXCLUDED.flo,
        pla = EXCLUDED.pla,
        acc = EXCLUDED.acc,
        all_flag = EXCLUDED.all_flag,
        boxes = EXCLUDED.boxes,
        dc_planning = EXCLUDED.dc_planning,
        dc_actual = EXCLUDED.dc_actual,
        dcs = EXCLUDED.dcs,
        dco = EXCLUDED.dco,
        cctag = EXCLUDED.cctag,
        vk = EXCLUDED.vk,
        pal = EXCLUDED.pal,
        raw = EXCLUDED.raw,
        updated_at = now()
    `,
    [
      row.id,
      row.action_date,
      row.country,
      row.code,
      row.group_with || "",
      row.carrier1_name || "",
      row.carrier2_name || "",
      row.matched_customer_name || "",
      row.matched_connect_name || "",
      row.matched_customer_code || "",
      row.matched_by || "",
      row.hours || "",
      row.flo === null || row.flo === undefined ? null : Boolean(row.flo),
      row.pla === null || row.pla === undefined ? null : Boolean(row.pla),
      row.acc === null || row.acc === undefined ? null : Boolean(row.acc),
      row.all_flag === null || row.all_flag === undefined ? null : Boolean(row.all_flag),
      numberOrNull(row.boxes),
      numberOrNull(row.dc_planning),
      numberOrNull(row.dc_actual),
      numberOrNull(row.dcs),
      numberOrNull(row.dco),
      numberOrNull(row.cctag),
      numberOrNull(row.vk),
      numberOrNull(row.pal),
      JSON.stringify(row.raw || {}),
    ],
  );
}

// Removes the rows for one date that the Fust API no longer returns (e.g.
// test data deleted in the Fust Planning app) -- keepIds is exactly the set
// the API just returned and saved for that date. Returns what was removed.
export async function deleteFustReferenceActionsForDateExcept(date, keepIds) {
  if (!pool || !date) {
    return [];
  }
  const result = await pool.query(
    `
      DELETE FROM fust_reference_actions
      WHERE action_date = $1::date AND NOT (id = ANY($2::text[]))
      RETURNING id, code, matched_customer_name
    `,
    [date, Array.isArray(keepIds) ? keepIds : []],
  );
  return result.rows;
}

// Feeds the Fust Planning viewer (the reference-code-level API import had no
// read path at all until this was added -- see saveFustReferenceAction above).
export async function getFustReferenceActions({ from, to } = {}) {
  if (!pool || !from || !to) {
    return [];
  }
  const result = await pool.query(
    `
      SELECT * FROM fust_reference_actions
      WHERE action_date >= $1::date AND action_date <= $2::date
      ORDER BY action_date DESC, code ASC
    `,
    [from, to],
  );
  return result.rows;
}

function mapUkdocsCsiParsedDocumentRow(row) {
  if (!row) {
    return null;
  }
  return {
    id: Number(row.id || 0),
    collection_id: String(row.collection_id || "").trim(),
    shipment_id: String(row.shipment_id || "").trim(),
    document_kind: String(row.document_kind || "").trim(),
    storage_name: String(row.storage_name || "").trim(),
    original_name: String(row.original_name || "").trim(),
    content_type: String(row.content_type || "").trim(),
    document_domain: String(row.document_domain || "").trim(),
    pcnu_number: String(row.pcnu_number || "").trim(),
    line_count: Number(row.line_count || 0),
    parsed_data: jsonValue(row.parsed_data_json, null),
    meta: jsonValue(row.meta_json, {}),
    parsed_at: row.parsed_at ? new Date(row.parsed_at).toISOString() : "",
    created_at: row.created_at ? new Date(row.created_at).toISOString() : "",
    updated_at: row.updated_at ? new Date(row.updated_at).toISOString() : "",
  };
}

function mapUkdocsCsiParsedDocumentItemRow(row) {
  return {
    row_index: Number(row?.row_index || 0),
    source_kind: String(row?.source_kind || "").trim(),
    source_document: String(row?.source_document || "").trim(),
    product_name: String(row?.product_name || "").trim(),
    botanical_name: String(row?.botanical_name || "").trim(),
    mapped_group: String(row?.mapped_group || "").trim(),
    commodity_code: String(row?.commodity_code || "").trim(),
    quantity: numberOrNull(row?.quantity),
    package_count: numberOrNull(row?.package_count),
    package_unit: String(row?.package_unit || "").trim(),
    quantity_unit: String(row?.quantity_unit || "").trim(),
    pcnu_number: String(row?.pcnu_number || "").trim(),
    raw_row: jsonValue(row?.raw_row_json, {}),
    created_at: row?.created_at ? new Date(row.created_at).toISOString() : "",
    updated_at: row?.updated_at ? new Date(row.updated_at).toISOString() : "",
  };
}

export async function saveUkdocsCsiParsedDocumentToDatabase(snapshot) {
  if (!pool || !snapshot?.collection_id || !snapshot?.document_kind || !snapshot?.storage_name) {
    return null;
  }

  const client = await pool.connect();
  try {
    await client.query("BEGIN");
    const documentResult = await client.query(
      `
        INSERT INTO ukdocs_csi_parsed_documents (
          collection_id, shipment_id, document_kind, storage_name, original_name, content_type,
          document_domain, pcnu_number, line_count, parsed_data_json, meta_json, parsed_at, updated_at
        )
        VALUES (
          $1, $2, $3, $4, $5, $6,
          $7, $8, $9, $10::jsonb, $11::jsonb, now(), now()
        )
        ON CONFLICT (collection_id, document_kind, storage_name) DO UPDATE SET
          shipment_id = EXCLUDED.shipment_id,
          original_name = EXCLUDED.original_name,
          content_type = EXCLUDED.content_type,
          document_domain = EXCLUDED.document_domain,
          pcnu_number = EXCLUDED.pcnu_number,
          line_count = EXCLUDED.line_count,
          parsed_data_json = EXCLUDED.parsed_data_json,
          meta_json = EXCLUDED.meta_json,
          parsed_at = now(),
          updated_at = now()
        RETURNING id
      `,
      [
        String(snapshot.collection_id || "").trim(),
        String(snapshot.shipment_id || "").trim(),
        String(snapshot.document_kind || "").trim(),
        String(snapshot.storage_name || "").trim(),
        String(snapshot.original_name || "").trim(),
        String(snapshot.content_type || "").trim(),
        String(snapshot.document_domain || "").trim(),
        String(snapshot.pcnu_number || "").trim(),
        Number(snapshot.line_count || 0),
        JSON.stringify(snapshot.parsed_data || null),
        JSON.stringify(snapshot.meta || {}),
      ],
    );
    const documentId = Number(documentResult.rows?.[0]?.id || 0);

    await client.query(
      `
        DELETE FROM ukdocs_csi_parsed_document_rows
        WHERE document_id = $1
      `,
      [documentId],
    );

    const rows = Array.isArray(snapshot.rows) ? snapshot.rows : [];
    for (let index = 0; index < rows.length; index += 1) {
      const row = rows[index] || {};
      await client.query(
        `
          INSERT INTO ukdocs_csi_parsed_document_rows (
            document_id, row_index, source_kind, source_document, product_name, botanical_name,
            mapped_group, commodity_code, quantity, package_count, package_unit, quantity_unit,
            pcnu_number, raw_row_json, updated_at
          )
          VALUES (
            $1, $2, $3, $4, $5, $6,
            $7, $8, $9, $10, $11, $12,
            $13, $14::jsonb, now()
          )
        `,
        [
          documentId,
          Number(row.row_index ?? index),
          String(row.source_kind || "").trim(),
          String(row.source_document || "").trim(),
          String(row.product_name || "").trim(),
          String(row.botanical_name || "").trim(),
          String(row.mapped_group || "").trim(),
          String(row.commodity_code || "").trim(),
          numberOrNull(row.quantity),
          numberOrNull(row.package_count),
          String(row.package_unit || "").trim(),
          String(row.quantity_unit || "").trim(),
          String(row.pcnu_number || "").trim(),
          JSON.stringify(row.raw_row || {}),
        ],
      );
    }

    await client.query("COMMIT");
    return documentId;
  } catch (error) {
    await client.query("ROLLBACK");
    throw error;
  } finally {
    client.release();
  }
}

export async function getUkdocsCsiParsedDocumentFromDatabase(collectionId, documentKind, storageName) {
  if (!pool || !collectionId || !documentKind || !storageName) {
    return null;
  }

  const documentResult = await pool.query(
    `
      SELECT *
      FROM ukdocs_csi_parsed_documents
      WHERE collection_id = $1
        AND document_kind = $2
        AND storage_name = $3
      LIMIT 1
    `,
    [String(collectionId).trim(), String(documentKind).trim(), String(storageName).trim()],
  );

  const mappedDocument = mapUkdocsCsiParsedDocumentRow(documentResult.rows?.[0]);
  if (!mappedDocument) {
    return null;
  }

  const rowsResult = await pool.query(
    `
      SELECT *
      FROM ukdocs_csi_parsed_document_rows
      WHERE document_id = $1
      ORDER BY row_index ASC, id ASC
    `,
    [mappedDocument.id],
  );

  return {
    ...mappedDocument,
    rows: (rowsResult.rows || []).map(mapUkdocsCsiParsedDocumentItemRow),
  };
}

export async function deleteUkdocsCsiParsedDocumentFromDatabase(collectionId, documentKind, storageName) {
  if (!pool || !collectionId || !documentKind || !storageName) {
    return;
  }
  await pool.query(
    `
      DELETE FROM ukdocs_csi_parsed_documents
      WHERE collection_id = $1
        AND document_kind = $2
        AND storage_name = $3
    `,
    [String(collectionId).trim(), String(documentKind).trim(), String(storageName).trim()],
  );
}

export async function markFustActionDeletedInDatabase(actionId, deletedAt = null, deletedBy = "") {
  if (!pool || !actionId) {
    return;
  }
  await pool.query(
    `
      UPDATE fust_actions
      SET deleted = true, deleted_at = COALESCE($2, now()), deleted_by = COALESCE($3, deleted_by), updated_at = now()
      WHERE id = $1
    `,
    [actionId, deletedAt || null, deletedBy || null],
  );
}

function jsonOrEmpty(value) {
  if (value === null || value === undefined) {
    return {};
  }
  if (typeof value === "object") {
    return value;
  }
  try {
    return JSON.parse(value);
  } catch {
    return {};
  }
}

function fustActionRowFromDatabase(row, documentsByKind) {
  return {
    id: row.id,
    type: row.type,
    action_date: row.action_date instanceof Date ? row.action_date.toISOString().slice(0, 10) : String(row.action_date || ""),
    week: row.week === null || row.week === undefined ? null : Number(row.week),
    day_name: row.day_name || "",
    country: row.country || "",
    customer_name: row.customer_name || "",
    customer_code: row.customer_code || "",
    connect_name: row.connect_name || "",
    remark: row.remark || "",
    fustbon_reference: row.fustbon_reference || "",
    fustfactuur_reference: row.fustfactuur_reference || "",
    metrics: {
      dc: Number(row.dc || 0),
      cctag: Number(row.cctag || 0),
      dcs: Number(row.dcs || 0),
      dco: Number(row.dco || 0),
      pal: Number(row.pal || 0),
      vk: Number(row.vk || 0),
    },
    created_by: row.created_by || "",
    created_at: row.created_at ? new Date(row.created_at).toISOString() : "",
    confirmed_at: row.confirmed_at ? new Date(row.confirmed_at).toISOString() : "",
    confirmed_by: row.confirmed_by || "",
    import_source: jsonOrEmpty(row.import_source),
    deleted: row.deleted === true,
    deleted_at: row.deleted_at ? new Date(row.deleted_at).toISOString() : "",
    deleted_by: row.deleted_by || "",
    sheet_sync: jsonOrEmpty(row.sheet_sync),
    email_sync: jsonOrEmpty(row.email_sync),
    db_sync: jsonOrEmpty(row.db_sync),
    confirmation_reminder: jsonOrEmpty(row.confirmation_reminder),
    cmr: documentsByKind.cmr || {},
    fustbon: documentsByKind.fustbon || {},
  };
}

// Mirror image of saveFustActionToDatabase(): reconstructs full action records
// (including nested document status) straight from Postgres, so it can be
// used as the primary read source for the live action list instead of
// re-parsing the Retour/Uitgaand sheets.
// The LAN warehouse backend (serverBackend.js) pushes here on its own
// schedule -- upserts the single status snapshot and appends whatever new
// activity events came in since the last push.
export async function upsertWarehouseStatus(data) {
  if (!pool) {
    return;
  }
  await pool.query(
    `
      INSERT INTO warehouse_status (id, data, updated_at)
      VALUES (1, $1::jsonb, now())
      ON CONFLICT (id) DO UPDATE SET data = $1::jsonb, updated_at = now()
    `,
    [JSON.stringify(data ?? {})],
  );
}

function warehouseEventTimestamp(event) {
  const parsed = new Date(event?.timestamp);
  return Number.isNaN(parsed.getTime()) ? new Date() : parsed;
}

export async function insertWarehouseActivityEvents(events) {
  if (!pool || !Array.isArray(events) || !events.length) {
    return;
  }
  const values = [];
  const placeholders = events.map((event) => {
    values.push(JSON.stringify(event ?? {}), warehouseEventTimestamp(event));
    return `($${values.length - 1}::jsonb, $${values.length})`;
  });
  await pool.query(
    `INSERT INTO warehouse_activity_log (event, ts) VALUES ${placeholders.join(", ")}`,
    values,
  );
}

export async function getWarehouseStatus() {
  if (!pool) {
    return {};
  }
  const result = await pool.query("SELECT data FROM warehouse_status WHERE id = 1");
  return result.rows[0]?.data || {};
}

export async function getWarehouseActivityLog({ date, limit } = {}) {
  if (!pool) {
    return [];
  }
  const safeLimit = Number(limit) > 0 ? Math.min(Number(limit), 20000) : 5000;
  const result = date
    ? await pool.query(
      "SELECT event FROM warehouse_activity_log WHERE ts::date = $1 ORDER BY ts DESC LIMIT $2",
      [date, safeLimit],
    )
    : await pool.query(
      "SELECT event FROM warehouse_activity_log ORDER BY ts DESC LIMIT $1",
      [safeLimit],
    );
  return result.rows.map((row) => row.event).reverse();
}

export async function getFustActionsFromDatabase() {
  if (!pool) {
    return [];
  }

  const [actionsResult, documentsResult] = await Promise.all([
    pool.query("SELECT * FROM fust_actions"),
    pool.query("SELECT * FROM fust_action_documents"),
  ]);

  const documentsByAction = new Map();
  for (const docRow of documentsResult.rows) {
    if (!documentsByAction.has(docRow.action_id)) {
      documentsByAction.set(docRow.action_id, {});
    }
    documentsByAction.get(docRow.action_id)[docRow.document_kind] = {
      status: docRow.status || "missing",
      file_id: docRow.file_id || "",
      file_name: docRow.file_name || "",
      web_link: docRow.web_link || "",
      mime_type: docRow.mime_type || "",
      folder_id: docRow.folder_id || "",
      error: docRow.error || "",
      uploaded_at: docRow.uploaded_at ? new Date(docRow.uploaded_at).toISOString() : "",
      uploaded_by: docRow.uploaded_by || "",
    };
  }

  return actionsResult.rows.map((row) => fustActionRowFromDatabase(row, documentsByAction.get(row.id) || {}));
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

    // job_types is an explicit opt-in allowlist, sent only by a poller that
    // knows how to handle that specific type (e.g. the dedicated shelf-count
    // poller sends ["shelf_count"]) -- the existing poller was deliberately
    // left untouched and never sends this, so its own claiming behavior for
    // every job type it already knows is completely unaffected. The one
    // exception: "shelf_count" is never handed to a poller that didn't ask
    // for it by name, so an unrelated/older poller can never claim (and
    // wrongly fail) a job type it has no idea how to run. Same for
    // "king_export" (the King ERP delivery poller, king-poller-app).
    const allowedJobTypes = Array.isArray(options.jobTypes) ? options.jobTypes.filter(Boolean) : null;
    const jobResult = await client.query(
      `
        SELECT *
        FROM llm_jobs
        WHERE status = 'pending'
          AND attempt_count < max_attempts
          AND (
            ($1::text[] IS NOT NULL AND job_type = ANY($1::text[]))
            OR ($1::text[] IS NULL AND job_type NOT IN ('shelf_count', 'king_export'))
          )
        ORDER BY priority DESC, created_at ASC
        LIMIT 1
        FOR UPDATE SKIP LOCKED
      `,
      [allowedJobTypes],
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

// getLlmQueueSnapshot() caps at the 50 most recently created jobs across
// every job type, for the admin UI. That cap makes it useless for finding a
// specific stale job once enough other jobs (any type) have been created
// since -- an old excel_to_pdf job can silently fall out of that window
// while still stuck pending/claimed. This query is scoped to one job_type
// and only pending/claimed rows, so it stays small and complete regardless
// of total queue traffic.
// Finished jobs keep their whole payload/result forever otherwise --
// shelf_count photos (~5-10 MB per reference, every night), CSI audit
// documents, invoice workbooks, the generated invoice PDF (already saved as
// a file by then) and King export PDFs. That grew the database by hundreds
// of MB a week (suspected cause of Postgres crashing into "recovery mode").
// Strips just those file fields from jobs that finished more than
// `olderThanHours` ago; the job rows and every other field stay.
export async function compactFinishedLlmJobs(olderThanHours = 6) {
  if (!pool) {
    return 0;
  }
  const result = await pool.query(
    `
      UPDATE llm_jobs
      SET payload_json = payload_json - 'vision_documents' - 'workbook_content_base64' - 'pdfs',
          result_json = CASE
            WHEN result_json->'excel_pdf_result' ? 'content_base64'
              THEN jsonb_set(result_json, '{excel_pdf_result}', (result_json->'excel_pdf_result') - 'content_base64')
            ELSE result_json
          END
      WHERE status IN ('done', 'failed')
        AND updated_at < now() - make_interval(hours => $1)
        AND (
          payload_json ?| ARRAY['vision_documents', 'workbook_content_base64', 'pdfs']
          OR result_json->'excel_pdf_result' ? 'content_base64'
        )
    `,
    [Math.max(1, Number(olderThanHours) || 6)],
  );
  return result.rowCount || 0;
}

export async function getActiveLlmJobsByType(jobType) {
  if (!pool) {
    return [];
  }
  const result = await pool.query(
    "SELECT * FROM llm_jobs WHERE job_type = $1 AND status IN ('pending', 'claimed') ORDER BY created_at ASC",
    [String(jobType || "").trim()],
  );
  return (result.rows || []).map(mapLlmJobRow);
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
    ALTER TABLE IF EXISTS fust_actions
    ADD COLUMN IF NOT EXISTS deleted_at timestamptz
  `,
  `
    ALTER TABLE IF EXISTS fust_actions
    ADD COLUMN IF NOT EXISTS deleted_by text
  `,
  `
    ALTER TABLE IF EXISTS fust_actions
    ADD COLUMN IF NOT EXISTS sheet_sync jsonb NOT NULL DEFAULT '{}'::jsonb
  `,
  `
    ALTER TABLE IF EXISTS fust_actions
    ADD COLUMN IF NOT EXISTS email_sync jsonb NOT NULL DEFAULT '{}'::jsonb
  `,
  `
    ALTER TABLE IF EXISTS fust_actions
    ADD COLUMN IF NOT EXISTS db_sync jsonb NOT NULL DEFAULT '{}'::jsonb
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
  `
    CREATE TABLE IF NOT EXISTS ukdocs_csi_parsed_documents (
      id bigserial PRIMARY KEY,
      collection_id text NOT NULL,
      shipment_id text NOT NULL DEFAULT '',
      document_kind text NOT NULL,
      storage_name text NOT NULL,
      original_name text,
      content_type text,
      document_domain text,
      pcnu_number text,
      line_count integer NOT NULL DEFAULT 0,
      parsed_data_json jsonb NOT NULL DEFAULT 'null'::jsonb,
      meta_json jsonb NOT NULL DEFAULT '{}'::jsonb,
      parsed_at timestamptz NOT NULL DEFAULT now(),
      created_at timestamptz NOT NULL DEFAULT now(),
      updated_at timestamptz NOT NULL DEFAULT now(),
      UNIQUE(collection_id, document_kind, storage_name)
    )
  `,
  `
    CREATE TABLE IF NOT EXISTS ukdocs_csi_parsed_document_rows (
      id bigserial PRIMARY KEY,
      document_id bigint NOT NULL REFERENCES ukdocs_csi_parsed_documents(id) ON DELETE CASCADE,
      row_index integer NOT NULL DEFAULT 0,
      source_kind text,
      source_document text,
      product_name text,
      botanical_name text,
      mapped_group text,
      commodity_code text,
      quantity numeric,
      package_count numeric,
      package_unit text,
      quantity_unit text,
      pcnu_number text,
      raw_row_json jsonb NOT NULL DEFAULT '{}'::jsonb,
      created_at timestamptz NOT NULL DEFAULT now(),
      updated_at timestamptz NOT NULL DEFAULT now()
    )
  `,
  `
    CREATE INDEX IF NOT EXISTS ukdocs_csi_parsed_documents_lookup_idx
    ON ukdocs_csi_parsed_documents (collection_id, document_kind, storage_name)
  `,
  `
    CREATE INDEX IF NOT EXISTS ukdocs_csi_parsed_document_rows_document_idx
    ON ukdocs_csi_parsed_document_rows (document_id, row_index)
  `,
  // Single always-id=1 row (upserted) -- the warehouse dashboard only ever
  // needs the current snapshot, not per-location history.
  `
    CREATE TABLE IF NOT EXISTS warehouse_status (
      id integer PRIMARY KEY DEFAULT 1,
      data jsonb NOT NULL,
      updated_at timestamptz NOT NULL DEFAULT now(),
      CHECK (id = 1)
    )
  `,
  `
    CREATE TABLE IF NOT EXISTS warehouse_activity_log (
      id bigserial PRIMARY KEY,
      event jsonb NOT NULL,
      ts timestamptz NOT NULL
    )
  `,
  `
    CREATE INDEX IF NOT EXISTS warehouse_activity_log_ts_idx
    ON warehouse_activity_log (ts)
  `,
  `
    CREATE TABLE IF NOT EXISTS fust_reference_actions (
      id text PRIMARY KEY,
      action_date date NOT NULL,
      country text NOT NULL,
      code text NOT NULL,
      group_with text,
      carrier1_name text,
      carrier2_name text,
      matched_customer_name text,
      matched_connect_name text,
      matched_customer_code text,
      matched_by text,
      hours text,
      flo boolean,
      pla boolean,
      acc boolean,
      all_flag boolean,
      boxes integer,
      dc_planning integer,
      dc_actual integer,
      dcs integer,
      dco integer,
      cctag integer,
      vk integer,
      pal integer,
      raw jsonb NOT NULL,
      imported_at timestamptz NOT NULL DEFAULT now(),
      updated_at timestamptz NOT NULL DEFAULT now()
    )
  `,
  `
    CREATE INDEX IF NOT EXISTS fust_reference_actions_date_idx
    ON fust_reference_actions (action_date)
  `,
  `
    CREATE TABLE IF NOT EXISTS inkoop_erp_lines (
      lot bigint PRIMARY KEY,
      erp_date date,
      pav text,
      suppl text,
      transp text,
      avc text,
      description text,
      pieces numeric,
      price numeric,
      t_price numeric,
      deb_no text,
      inv_no text,
      raw jsonb NOT NULL,
      imported_at timestamptz NOT NULL DEFAULT now(),
      updated_at timestamptz NOT NULL DEFAULT now()
    )
  `,
  `
    CREATE INDEX IF NOT EXISTS inkoop_erp_lines_pav_idx ON inkoop_erp_lines (pav)
  `,
  `
    CREATE INDEX IF NOT EXISTS inkoop_erp_lines_suppl_idx ON inkoop_erp_lines (suppl)
  `,
  `
    CREATE INDEX IF NOT EXISTS inkoop_erp_lines_date_idx ON inkoop_erp_lines (erp_date)
  `,
  `
    CREATE TABLE IF NOT EXISTS inkoop_invoice_lines (
      id text PRIMARY KEY,
      invoice_number text NOT NULL,
      invoice_type text NOT NULL,
      reference_bt text,
      supplier_gln text,
      supplier_fh_number text,
      supplier_name text,
      description text,
      quantity numeric,
      unit_price numeric,
      total numeric,
      raw jsonb NOT NULL,
      source_file_name text,
      invoice_date date,
      imported_at timestamptz NOT NULL DEFAULT now(),
      updated_at timestamptz NOT NULL DEFAULT now()
    )
  `,
  `
    CREATE INDEX IF NOT EXISTS inkoop_invoice_lines_reference_bt_idx ON inkoop_invoice_lines (reference_bt)
  `,
  `
    CREATE INDEX IF NOT EXISTS inkoop_invoice_lines_supplier_name_idx ON inkoop_invoice_lines (supplier_name)
  `,
  // The invoice-number prefix (e.g. "063155") identifying which of our own
  // buying entities an invoice was billed to -- verified 1:1 against the
  // real InvoiceeParty name in the XML (company_name), so both ride
  // together as a single filterable dimension across the whole page.
  `
    ALTER TABLE inkoop_invoice_lines ADD COLUMN IF NOT EXISTS company_number text
  `,
  `
    ALTER TABLE inkoop_invoice_lines ADD COLUMN IF NOT EXISTS company_name text
  `,
  `
    CREATE INDEX IF NOT EXISTS inkoop_invoice_lines_company_number_idx ON inkoop_invoice_lines (company_number)
  `,
  `
    CREATE TABLE IF NOT EXISTS inkoop_issue_status (
      id text PRIMARY KEY,
      issue_type text NOT NULL,
      status text NOT NULL DEFAULT 'open',
      assigned_to text,
      note text,
      updated_by text,
      updated_at timestamptz NOT NULL DEFAULT now()
    )
  `,
  // One row per ERP purchase lot, carrying the "Dispatched to : <customer
  // code>" rows the ERP "Screen" export now interleaves after each lot --
  // where that lot's stock actually went. A separate, independently
  // scheduled upload from the day-to-day Klokfactuur/Connect/Handel
  // reconciliation ledger (inkoop_erp_lines) -- keyed the same way (lot) but
  // deliberately its own table, since it's sourced from its own dump file on
  // its own cadence, not tied to a specific invoice-compare run.
  `
    CREATE TABLE IF NOT EXISTS inkoop_dispatch_lots (
      lot bigint PRIMARY KEY,
      erp_date date,
      pav text,
      suppl text,
      description text,
      pieces numeric,
      price numeric,
      t_price numeric,
      dispatches jsonb NOT NULL DEFAULT '[]'::jsonb,
      imported_at timestamptz NOT NULL DEFAULT now(),
      updated_at timestamptz NOT NULL DEFAULT now()
    )
  `,
  `
    CREATE INDEX IF NOT EXISTS inkoop_dispatch_lots_date_idx ON inkoop_dispatch_lots (erp_date)
  `,
  // One row per FloraHolland invoice: the totals King's journal export
  // needs (total incl. BTW, BTW per category, klok location) -- parsed from
  // the invoice XML header (see inkoop_veiling_worker.py
  // parse_invoice_header), which the per-line ledger doesn't carry.
  `
    CREATE TABLE IF NOT EXISTS inkoop_invoice_headers (
      invoice_number text PRIMARY KEY,
      invoice_type text,
      company_number text,
      company_name text,
      invoice_date date,
      grand_total numeric,
      currency text,
      vat_subtotals jsonb NOT NULL DEFAULT '[]'::jsonb,
      klok_location text,
      klok_locations jsonb NOT NULL DEFAULT '[]'::jsonb,
      source_file_name text,
      imported_at timestamptz NOT NULL DEFAULT now(),
      updated_at timestamptz NOT NULL DEFAULT now()
    )
  `,
  `
    CREATE INDEX IF NOT EXISTS inkoop_invoice_headers_date_idx ON inkoop_invoice_headers (invoice_date)
  `,
  // What has been sent to King, per invoice -- the single source of truth
  // that keeps an invoice from being booked twice. Deliberately NOT cleared
  // by "Start fresh" (clearInkoopUploads): re-uploading an invoice later
  // must still know it already went to King.
  `
    CREATE TABLE IF NOT EXISTS king_exports (
      invoice_number text PRIMARY KEY,
      batch_id text NOT NULL,
      stuknummer integer,
      extern_id text,
      status text NOT NULL DEFAULT 'queued',
      job_id text,
      error_text text,
      exported_by text,
      exported_at timestamptz NOT NULL DEFAULT now(),
      delivered_at timestamptz,
      journal_json jsonb
    )
  `,
  `
    CREATE INDEX IF NOT EXISTS king_exports_batch_idx ON king_exports (batch_id)
  `,
  // Corrected design (superseding a first attempt that invented a new
  // per-trolley identity/ingest pipeline that turned out to be unnecessary):
  // warehouse_activity_log's own "scan_complete" events already carry
  // photoCount/rfidCount/trolleyCount per customer reference -- once
  // photoCount reaches trolleyCount, that reference's photos are complete
  // and ready to count, no new "trolley" concept or ingest route needed.
  // shelf_counts is therefore keyed by (customer_reference, date), not a
  // manufactured trolley id.
  //
  // This used to DROP TABLE shelf_counts/warehouse_trolley_scans here first,
  // to clear out the original attempt's shape (trolley_scan_id PRIMARY KEY),
  // which never held any real data since its RFID-portal ingest route was
  // never wired up externally. That was only ever meant to run once --
  // applyDatabaseMigrations re-runs this whole list on every server start
  // (confirmed real, and confirmed the actual bug: it silently wiped every
  // real shelf_counts row -- a full night's worth of LLM-counted
  // shelves/levels/confidence -- on every single deploy, forever, since an
  // unconditional DROP has no way to know the one-time cleanup it was meant
  // for already happened). Removed now that the old shape is long gone and
  // shelf_counts holds real production data that must survive a restart.
  // expected_average/deviation stay null in Phase 1 -- there is no manual
  // "expected count" config; a later training/derivation phase fills these
  // in from accumulated real counts, not a config screen.
  `
    CREATE TABLE IF NOT EXISTS shelf_counts (
      customer_reference text NOT NULL,
      nightly_run_date date NOT NULL,
      drive_folder_name text NOT NULL DEFAULT '',
      trolley_count integer NOT NULL DEFAULT 0,
      photo_count integer NOT NULL DEFAULT 0,
      shelf_count integer,
      level_count integer,
      confidence numeric,
      expected_average numeric,
      deviation numeric,
      extension_count integer,
      extension_expected numeric,
      extension_deviation numeric,
      status text NOT NULL DEFAULT 'pending',
      model_version text NOT NULL DEFAULT '',
      job_id text,
      error_text text NOT NULL DEFAULT '',
      processed_at timestamptz,
      created_at timestamptz NOT NULL DEFAULT now(),
      updated_at timestamptz NOT NULL DEFAULT now(),
      PRIMARY KEY (customer_reference, nightly_run_date)
    )
  `,
  `
    CREATE INDEX IF NOT EXISTS shelf_counts_run_date_idx ON shelf_counts (nightly_run_date)
  `,
  // Trolley pole extensions (the "DCO" code) -- 0-4 per trolley, added when
  // a trolley is loaded high enough to need them. Counted by the same LLM
  // job as shelf_count/level_count, shadowed against Fust Planning's DCO
  // the same way shelf_count is shadowed against DC.
  `
    ALTER TABLE shelf_counts ADD COLUMN IF NOT EXISTS extension_count integer
  `,
  `
    ALTER TABLE shelf_counts ADD COLUMN IF NOT EXISTS extension_expected numeric
  `,
  `
    ALTER TABLE shelf_counts ADD COLUMN IF NOT EXISTS extension_deviation numeric
  `,
  // A reference can have several trolleys (trolley_count) -- the model is
  // told how many and returns one {shelves, levels, extensions} entry per
  // trolley next to the totals, so a multi-trolley count can be checked.
  `
    ALTER TABLE shelf_counts ADD COLUMN IF NOT EXISTS per_trolley jsonb
  `,
  // Side-by-side count from the trained YOLO shelf model (shelf-training,
  // run by shelf-poller-app next to the Ollama count) -- for comparison
  // only; shelf_count/level_count stay the official Ollama figures.
  `
    ALTER TABLE shelf_counts ADD COLUMN IF NOT EXISTS trained_model jsonb
  `,
  // Fust Planning's trolley count (DC) for the reference, next to the RFID
  // scan's own trolley_count -- expected_average holds the shelves (DCS).
  `
    ALTER TABLE shelf_counts ADD COLUMN IF NOT EXISTS expected_trolleys numeric
  `,
  // "Combined"/"Mixed" references (Halindeling app maincode) are loaded onto
  // their main code's trolleys and never photographed themselves -- the
  // main code they belong to.
  `
    ALTER TABLE shelf_counts ADD COLUMN IF NOT EXISTS combined_into text
  `,
  // One row per calendar day the nightly trigger has run for -- its own
  // existence for a given run_date is what makes the trigger idempotent
  // (skip if already run) while still being explicitly re-runnable (a
  // manual re-run for a specific date just calls the same function again).
  `
    CREATE TABLE IF NOT EXISTS shelf_count_nightly_runs (
      run_date date PRIMARY KEY,
      status text NOT NULL DEFAULT 'running',
      totals jsonb NOT NULL DEFAULT '{}'::jsonb,
      issues jsonb NOT NULL DEFAULT '[]'::jsonb,
      started_at timestamptz NOT NULL DEFAULT now(),
      completed_at timestamptz,
      created_at timestamptz NOT NULL DEFAULT now(),
      updated_at timestamptz NOT NULL DEFAULT now()
    )
  `,
];

// Every statement in databaseMigrations is idempotent (CREATE TABLE IF NOT
// EXISTS / ADD COLUMN IF NOT EXISTS / CREATE INDEX IF NOT EXISTS), so this is
// always safe to call again against an already-initialized pool -- notably
// after the standby PC restores a Postgres dump pulled from Render, which
// would otherwise leave the schema however that dump happened to look,
// until the next full server restart re-ran migrations.
export async function applyDatabaseMigrations() {
  if (!pool) {
    return { ok: false, error: "Database pool is not initialized" };
  }
  try {
    for (const migration of databaseMigrations) {
      await pool.query(migration);
    }
    return { ok: true, error: "" };
  } catch (error) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) };
  }
}

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
    const migrationResult = await applyDatabaseMigrations();
    if (!migrationResult.ok) {
      throw new Error(migrationResult.error);
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

// --- Standby backup snapshot, without pg_dump ---
// Render's pg_dump (15) refuses its own newer Postgres (18), and a native
// Render service can't install another version. So the standby copy is made
// through the normal connection instead: every table's rows as JSON lines,
// loaded on the office PC into the same tables (its own migrations create
// them) -- independent of either side's Postgres version.
const SNAPSHOT_BATCH_ROWS = 500;
const SNAPSHOT_BATCH_BYTES = 8 * 1024 * 1024;

const quoteIdent = (name) => `"${String(name).replace(/"/g, '""')}"`;

async function listPublicTablesParentsFirst(client) {
  const tables = (await client.query(
    `SELECT table_name FROM information_schema.tables WHERE table_schema = 'public' AND table_type = 'BASE TABLE' ORDER BY table_name`,
  )).rows.map((row) => row.table_name);
  const deps = (await client.query(`
    SELECT child.relname AS child, parent.relname AS parent
    FROM pg_constraint c
    JOIN pg_class child ON child.oid = c.conrelid
    JOIN pg_class parent ON parent.oid = c.confrelid
    JOIN pg_namespace n ON n.oid = child.relnamespace
    WHERE c.contype = 'f' AND n.nspname = 'public'
  `)).rows;
  // Parents before children, so foreign keys hold while loading.
  const ordered = [];
  const visit = (table, seen = new Set()) => {
    if (ordered.includes(table) || seen.has(table)) return;
    seen.add(table);
    for (const dep of deps) {
      if (dep.child === table && dep.parent !== table) visit(dep.parent, seen);
    }
    ordered.push(table);
  };
  tables.forEach((table) => visit(table));
  return ordered.filter((table) => tables.includes(table));
}

// Newline-delimited JSON: a header line, then per table a {"table"} line
// followed by one line per row -- handed line by line to `writeLine` (which
// may return a promise, for backpressure), never built up as one string: the
// whole database as one string ran Render out of memory (Oct 2026).
// Finished LLM jobs go without their payload/result (photos, PDFs, model
// output): a standby never re-runs those, and they're most of the size.
const SNAPSHOT_FETCH_ROWS = 200;

function snapshotRowSelect(table) {
  if (table === "llm_jobs") {
    return `SELECT (to_jsonb(t) || CASE WHEN t.finished_at IS NOT NULL THEN '{"payload_json":{},"result_json":{}}'::jsonb ELSE '{}'::jsonb END)::text AS row FROM llm_jobs t`;
  }
  return `SELECT to_jsonb(t)::text AS row FROM ${quoteIdent(table)} t`;
}

export async function exportDatabaseSnapshot(writeLine) {
  if (!pool) {
    throw new Error("Database is not initialized");
  }
  const client = await pool.connect();
  try {
    // One consistent point in time across all tables.
    await client.query("BEGIN ISOLATION LEVEL REPEATABLE READ READ ONLY");
    const tables = await listPublicTablesParentsFirst(client);
    await writeLine(JSON.stringify({ format: "snappysjaak-db-snapshot", version: 1, created_at: new Date().toISOString(), tables }));
    for (const table of tables) {
      await writeLine(JSON.stringify({ table }));
      for (let offset = 0; ; offset += SNAPSHOT_FETCH_ROWS) {
        const { rows } = await client.query(`${snapshotRowSelect(table)} ORDER BY ctid LIMIT ${SNAPSHOT_FETCH_ROWS} OFFSET ${offset}`);
        for (const row of rows) {
          await writeLine(row.row);
        }
        if (rows.length < SNAPSHOT_FETCH_ROWS) break;
      }
    }
    await client.query("COMMIT");
  } catch (error) {
    await client.query("ROLLBACK").catch(() => {});
    throw error;
  } finally {
    client.release();
  }
}

// Replaces the local tables' contents with the snapshot, all in one
// transaction: either everything is the new copy or nothing changed.
// `lines` is an (async) iterable of the snapshot's lines, read one at a time
// so the whole copy never sits in memory as one string either.
export async function importDatabaseSnapshot(lines) {
  if (!pool) {
    throw new Error("Database is not initialized");
  }
  const iterator = lines[Symbol.asyncIterator] ? lines[Symbol.asyncIterator]() : lines[Symbol.iterator]();
  let first = await iterator.next();
  while (!first.done && !String(first.value).trim()) first = await iterator.next();
  const header = JSON.parse(first.done ? "{}" : first.value);
  if (header.format !== "snappysjaak-db-snapshot") {
    throw new Error("Not a SnappySjaak database snapshot");
  }
  const client = await pool.connect();
  const counts = {};
  try {
    await client.query("BEGIN");
    const localTables = new Set((await client.query(
      `SELECT table_name FROM information_schema.tables WHERE table_schema = 'public' AND table_type = 'BASE TABLE'`,
    )).rows.map((row) => row.table_name));
    const tables = (header.tables || []).filter((table) => localTables.has(table));
    if (tables.length) {
      await client.query(`TRUNCATE ${tables.map(quoteIdent).join(", ")} RESTART IDENTITY CASCADE`);
    }
    const columnsOf = async (table) => (await client.query(
      `SELECT column_name FROM information_schema.columns WHERE table_schema = 'public' AND table_name = $1 AND is_generated = 'NEVER'`,
      [table],
    )).rows.map((row) => row.column_name);

    let table = null;
    let columns = [];
    let batch = [];
    let batchBytes = 0;
    const flush = async () => {
      if (!table || !batch.length) return;
      // Only columns present on both sides; a column only this PC has keeps
      // its default.
      const present = columns.filter((column) => Object.prototype.hasOwnProperty.call(batch[0], column));
      const list = present.map(quoteIdent).join(", ");
      await client.query(
        `INSERT INTO ${quoteIdent(table)} (${list}) SELECT ${list} FROM jsonb_populate_recordset(NULL::${quoteIdent(table)}, $1::jsonb)`,
        [JSON.stringify(batch)],
      );
      counts[table] = (counts[table] || 0) + batch.length;
      batch = [];
      batchBytes = 0;
    };
    for (let next = await iterator.next(); !next.done; next = await iterator.next()) {
      const line = next.value;
      if (!line || !String(line).trim()) continue;
      const parsed = JSON.parse(line);
      if (parsed && typeof parsed.table === "string" && Object.keys(parsed).length === 1) {
        await flush();
        table = localTables.has(parsed.table) ? parsed.table : null;
        columns = table ? await columnsOf(table) : [];
        if (table) counts[table] = 0;
        continue;
      }
      if (!table) continue;
      batch.push(parsed);
      batchBytes += line.length;
      if (batch.length >= SNAPSHOT_BATCH_ROWS || batchBytes >= SNAPSHOT_BATCH_BYTES) {
        await flush();
      }
    }
    await flush();

    // Serial ids continue after the highest copied id.
    const sequences = (await client.query(`
      SELECT table_name, column_name FROM information_schema.columns
      WHERE table_schema = 'public' AND column_default LIKE 'nextval(%'
    `)).rows;
    for (const { table_name: seqTable, column_name: seqColumn } of sequences) {
      await client.query(
        `SELECT setval(pg_get_serial_sequence($1, $2), COALESCE((SELECT MAX(${quoteIdent(seqColumn)}) FROM ${quoteIdent(seqTable)}), 0) + 1, false)`,
        [`public.${quoteIdent(seqTable)}`, seqColumn],
      );
    }
    await client.query("COMMIT");
  } catch (error) {
    await client.query("ROLLBACK").catch(() => {});
    throw error;
  } finally {
    client.release();
  }
  return { tables: counts, snapshot_created_at: header.created_at || "" };
}
