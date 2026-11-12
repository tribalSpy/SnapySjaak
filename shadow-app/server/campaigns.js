// Campagnes: the monthly offer catalogue (about 50 articles for next month's
// offer), the awareness campaign that runs before it, and the agreements
// with the growers about those articles.
//
// Articles are built from our own purchase invoices (VBN code, invoice name,
// grower, packages x per package, price) and kept as reusable master cards
// with their steady conditions. Everything lives in its own camp_* tables
// (server/db.js); invoice data is only read. Light work only -- media,
// publishing and other heavy work go to the BEAST PC through the job queue.
import { dbQuery, isDatabaseEnabled, searchInkoopArticles, searchInkoopProductCards } from "./db.js";
import { createSystemTask } from "./tasks.js";

export const CAMPAIGN_STAGES = ["selecting", "growers", "content", "running", "offer", "evaluating", "closed", "cancelled"];
export const CAMPAIGN_ITEM_STATUSES = ["shortlist", "selected", "deal_ok", "content_ok", "live", "evaluated", "dropped"];
export const GROWER_DEAL_STATUSES = ["idea", "contacted", "negotiating", "agreed", "confirmed", "delivering", "closed", "declined"];
export const GROWER_CONTACT_KINDS = ["note", "call", "email", "visit", "whatsapp", "meeting"];

// VBN feature codes we know the meaning of; the others are shown as code.
export const VBN_FEATURE_LABELS = {
  S20: "Length (cm)",
  S01: "Pot size (cm)",
  S02: "Plant height (cm)",
};

const CONDITION_KEYS = ["length_cm", "pot_cm", "height_cm", "stems_per_bunch", "packages", "per_package", "pieces", "unit",
  "fust", "quality", "colour", "origin", "ref_price", "min_order", "delivery_days"];

function text(value, max = 500) {
  return String(value ?? "").trim().slice(0, max);
}

function num(value) {
  if (value === null || value === undefined || value === "") return null;
  const parsed = Number(String(value).replace(",", "."));
  return Number.isFinite(parsed) ? parsed : null;
}

function isoDay(value) {
  const match = String(value || "").match(/^(\d{4})-(\d{2})-(\d{2})/);
  return match ? `${match[1]}-${match[2]}-${match[3]}` : null;
}

// A Postgres DATE comes back as a Date at local midnight: read it with the
// local calendar (toISOString would give the day before in Europe).
function dbDay(value) {
  if (value instanceof Date) {
    const pad = (number) => String(number).padStart(2, "0");
    return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())}`;
  }
  return value ? String(value).slice(0, 10) : null;
}

function localToday() {
  return dbDay(new Date());
}

function isMonth(value) {
  return /^\d{4}-(0[1-9]|1[0-2])$/.test(String(value || ""));
}

function monthStart(month, shift = 0) {
  const [year, mon] = month.split("-").map(Number);
  const date = new Date(Date.UTC(year, mon - 1 + shift, 1));
  return date.toISOString().slice(0, 10);
}

function addDays(day, days) {
  const date = new Date(`${day}T00:00:00Z`);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

// The default timeline of an offer month M: campaign window = month M-1;
// articles chosen 3 weeks before the window, growers agreed 1 week before,
// content ready when the window opens, evaluation 2 weeks into M+1.
export function defaultCampaignTimeline(offerMonth, leadMonths = 1) {
  const lead = Math.max(1, Math.min(3, Number(leadMonths) || 1));
  const start = monthStart(offerMonth, -lead);
  const end = addDays(monthStart(offerMonth, 0), -1);
  return {
    campaign_start: start,
    campaign_end: end,
    deadlines: {
      articles: addDays(start, -21),
      growers: addDays(start, -7),
      content: start,
      evaluation: addDays(monthStart(offerMonth, 1), 14),
    },
  };
}

function normalizeConditions(value) {
  const source = value && typeof value === "object" && !Array.isArray(value) ? value : {};
  const out = {};
  for (const key of CONDITION_KEYS) {
    if (source[key] === undefined || source[key] === null || source[key] === "") continue;
    out[key] = ["unit", "fust", "quality", "colour", "origin", "delivery_days"].includes(key) ? text(source[key], 120) : num(source[key]);
    if (out[key] === null) delete out[key];
  }
  // Pieces = packages x per package, unless typed by hand.
  if (out.pieces === undefined && out.packages && out.per_package) out.pieces = Math.round(out.packages * out.per_package * 100) / 100;
  if (source.features && typeof source.features === "object") {
    const features = {};
    for (const [code, featureValue] of Object.entries(source.features)) {
      const key = text(code, 10).toUpperCase();
      if (key) features[key] = text(featureValue, 40);
    }
    if (Object.keys(features).length) out.features = features;
  }
  if (source.extra && typeof source.extra === "object") {
    const extra = {};
    for (const [label, extraValue] of Object.entries(source.extra).slice(0, 30)) {
      const key = text(label, 60);
      if (key) extra[key] = text(extraValue, 200);
    }
    if (Object.keys(extra).length) out.extra = extra;
  }
  return out;
}

function normalizeGrowers(value) {
  return (Array.isArray(value) ? value : []).slice(0, 30).map((grower) => ({
    name: text(grower?.name, 160),
    code: text(grower?.code, 40),
    gln: text(grower?.gln, 20),
    fh: text(grower?.fh, 20),
    last_price: num(grower?.last_price),
    last_date: isoDay(grower?.last_date),
    pieces: num(grower?.pieces),
  })).filter((grower) => grower.name || grower.code || grower.gln);
}

// How sales are found back for this article (phase 6): VBN codes, grower
// codes and name terms. General words never count alone (the matcher side).
function normalizeMatchRules(value) {
  const source = value && typeof value === "object" ? value : {};
  const list = (items, max) => [...new Set((Array.isArray(items) ? items : []).map((item) => text(item, max)).filter(Boolean))].slice(0, 40);
  return { vbn_codes: list(source.vbn_codes, 20), grower_codes: list(source.grower_codes, 40), name_terms: list(source.name_terms, 60) };
}

function articleRow(row) {
  if (!row) return null;
  return {
    id: Number(row.id),
    name: row.name,
    name_fr: row.name_fr,
    name_en: row.name_en,
    vbn_codes: row.vbn_codes || [],
    product_id: row.product_id,
    hoofdgroep: row.hoofdgroep,
    sales_group: row.sales_group,
    conditions: row.conditions || {},
    growers: row.growers || [],
    match_rules: row.match_rules || {},
    photo_asset_ids: row.photo_asset_ids || [],
    notes: row.notes,
    archived: row.archived,
    created_by: row.created_by,
    created_at: row.created_at,
    updated_by: row.updated_by,
    updated_at: row.updated_at,
    campaigns: row.campaigns === undefined ? undefined : Number(row.campaigns) || 0,
  };
}

function campaignRow(row) {
  if (!row) return null;
  const day = dbDay;
  return {
    id: Number(row.id),
    offer_month: row.offer_month,
    title_fr: row.title_fr,
    title_en: row.title_en,
    campaign_start: day(row.campaign_start),
    campaign_end: day(row.campaign_end),
    stage: row.stage,
    deadlines: row.deadlines || {},
    notes: row.notes,
    created_by: row.created_by,
    created_at: row.created_at,
    updated_by: row.updated_by,
    updated_at: row.updated_at,
    counts: row.counts || undefined,
    deal_counts: row.deal_counts || undefined,
  };
}

function dealRow(row) {
  if (!row) return null;
  const day = dbDay;
  return {
    id: Number(row.id),
    item_id: Number(row.item_id),
    grower_name: row.grower_name,
    grower_code: row.grower_code,
    grower_gln: row.grower_gln,
    grower_fh: row.grower_fh,
    status: row.status,
    agreed_price: row.agreed_price === null || row.agreed_price === undefined ? null : Number(row.agreed_price),
    qty_plan: row.qty_plan || [],
    conditions: row.conditions || {},
    valid_from: day(row.valid_from),
    valid_to: day(row.valid_to),
    next_action_at: day(row.next_action_at),
    next_action: row.next_action,
    created_by: row.created_by,
    created_at: row.created_at,
    updated_by: row.updated_by,
    updated_at: row.updated_at,
    contacts: row.contacts === undefined ? undefined : Number(row.contacts) || 0,
    last_contact_at: row.last_contact_at || null,
    // Joined in lists:
    article_id: row.article_id === undefined ? undefined : Number(row.article_id),
    article_name: row.article_name,
    campaign_id: row.campaign_id === undefined ? undefined : Number(row.campaign_id),
    offer_month: row.offer_month,
  };
}

// ---------------------------------------------------------------------------
// Settings

function normalizeSettings(value) {
  const source = value && typeof value === "object" ? value : {};
  const target = text(source.task_target, 80);
  return {
    // Who gets the campaign tasks (deadlines, grower next actions):
    // "team:<key>" or "user:<name>"; empty = no tasks.
    task_target: /^(user|team):./.test(target) ? target : "",
    lead_months: Math.max(1, Math.min(3, Number(source.lead_months) || 1)),
    target_articles: Math.max(1, Math.min(500, Number(source.target_articles) || 50)),
    deadline_task_days: Math.max(0, Math.min(30, Number(source.deadline_task_days ?? 3))),
  };
}

export async function readCampaignSettings() {
  if (!isDatabaseEnabled()) return normalizeSettings(null);
  const result = await dbQuery("SELECT value FROM camp_settings WHERE key = 'config'");
  return normalizeSettings(result.rows[0]?.value);
}

async function writeCampaignSettings(value, by) {
  const normalized = normalizeSettings(value);
  await dbQuery(
    `INSERT INTO camp_settings (key, value, updated_by, updated_at) VALUES ('config', $1, $2, now())
     ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_by = EXCLUDED.updated_by, updated_at = now()`,
    [JSON.stringify(normalized), by || ""],
  );
  return normalized;
}

function taskAssignee(target) {
  if (target.startsWith("team:")) return { assignee_team: target.slice(5) };
  if (target.startsWith("user:")) return { assignee_user: target.slice(5) };
  return null;
}

// ---------------------------------------------------------------------------
// Tasks made by the module (best effort, never throws).

async function taskForDealNextAction(deal, articleName, offerMonth, by) {
  try {
    if (!deal?.next_action_at) return;
    const settings = await readCampaignSettings();
    const assignee = taskAssignee(settings.task_target);
    if (!assignee) return;
    await createSystemTask({
      ...assignee,
      title: `Grower ${deal.grower_name || deal.grower_code}: ${deal.next_action || "follow up"} (${articleName})`,
      body: `Campagne ${offerMonth}: ${articleName}\nGrower deal status: ${deal.status}${deal.agreed_price !== null ? `\nAgreed price: ${deal.agreed_price}` : ""}`,
      due_date: deal.next_action_at,
      priority: "normal",
      source: "campaign",
      source_key: `camp-deal-${deal.id}-${deal.next_action_at}`,
      link: { page: "campaigns", tab: "growers", deal_id: deal.id },
      created_by: by || "system",
    });
  } catch (error) {
    console.error("Campaign deal task not made:", error?.message || error);
  }
}

const DEADLINE_TITLES = {
  articles: "choose the articles",
  growers: "grower deals done",
  content: "content ready, campaign starts",
  evaluation: "evaluate the campaign",
};

// Daily: a task for every campaign deadline coming up within the setting's
// days (once per campaign + deadline + date).
export async function runCampaignDeadlineTasks() {
  if (!isDatabaseEnabled()) return { made: 0 };
  const settings = await readCampaignSettings();
  const assignee = taskAssignee(settings.task_target);
  if (!assignee) return { made: 0 };
  const result = await dbQuery("SELECT * FROM camp_campaigns WHERE stage NOT IN ('closed', 'cancelled')");
  const today = localToday();
  const until = addDays(today, settings.deadline_task_days);
  let made = 0;
  for (const campaign of result.rows.map(campaignRow)) {
    for (const [key, day] of Object.entries(campaign.deadlines || {})) {
      if (!isoDay(day) || day > until || day < addDays(today, -7)) continue;
      const row = await createSystemTask({
        ...assignee,
        title: `Campagne ${campaign.offer_month}: ${DEADLINE_TITLES[key] || key} (${day})`,
        body: `Offer month ${campaign.offer_month}, campaign ${campaign.campaign_start || "?"} - ${campaign.campaign_end || "?"}.`,
        due_date: day,
        priority: day <= today ? "high" : "normal",
        source: "campaign",
        source_key: `camp-${campaign.id}-${key}-${day}`,
        link: { page: "campaigns", campaign_id: campaign.id },
      });
      if (row) made += 1;
    }
  }
  return { made };
}

// ---------------------------------------------------------------------------
// Queries

async function listCampaigns() {
  const result = await dbQuery(`
    SELECT c.*,
      COALESCE((SELECT jsonb_object_agg(status, n) FROM (SELECT status, count(*)::int AS n FROM camp_items i WHERE i.campaign_id = c.id GROUP BY status) s), '{}'::jsonb) AS counts,
      COALESCE((SELECT jsonb_object_agg(status, n) FROM (SELECT d.status, count(*)::int AS n FROM camp_grower_deals d JOIN camp_items i ON i.id = d.item_id WHERE i.campaign_id = c.id GROUP BY d.status) s), '{}'::jsonb) AS deal_counts
    FROM camp_campaigns c ORDER BY c.offer_month DESC
  `);
  return result.rows.map(campaignRow);
}

async function getCampaign(id) {
  const result = await dbQuery("SELECT * FROM camp_campaigns WHERE id = $1", [id]);
  return campaignRow(result.rows[0]);
}

async function getCampaignItems(campaignId) {
  const result = await dbQuery(
    `SELECT i.id, i.campaign_id, i.article_id, i.status, i.offer_price::float8 AS offer_price, i.planned_qty::float8 AS planned_qty,
       i.sort_order, i.notes, i.updated_by, i.updated_at, to_jsonb(a.*) AS article,
       COALESCE((SELECT jsonb_agg(jsonb_build_object('id', d.id, 'grower_name', d.grower_name, 'grower_code', d.grower_code, 'status', d.status,
         'agreed_price', d.agreed_price::float8, 'next_action_at', d.next_action_at, 'next_action', d.next_action) ORDER BY d.id)
         FROM camp_grower_deals d WHERE d.item_id = i.id), '[]'::jsonb) AS deals
     FROM camp_items i JOIN camp_articles a ON a.id = i.article_id
     WHERE i.campaign_id = $1
     ORDER BY i.sort_order, i.id`,
    [campaignId],
  );
  return result.rows.map((row) => ({
    id: Number(row.id),
    campaign_id: Number(row.campaign_id),
    article_id: Number(row.article_id),
    status: row.status,
    offer_price: row.offer_price,
    planned_qty: row.planned_qty,
    sort_order: row.sort_order,
    notes: row.notes,
    updated_by: row.updated_by,
    updated_at: row.updated_at,
    article: articleRow(row.article),
    deals: row.deals || [],
  }));
}

async function listArticles({ q, includeArchived = false, limit = 200 } = {}) {
  const values = [];
  const conditions = [];
  if (!includeArchived) conditions.push("NOT a.archived");
  for (const word of String(q || "").split(/\s+/).filter(Boolean).slice(0, 6)) {
    values.push(`%${word.replace(/[\\%_]/g, (match) => `\\${match}`)}%`);
    const p = `$${values.length}`;
    conditions.push(`(a.name ILIKE ${p} OR a.name_fr ILIKE ${p} OR a.name_en ILIKE ${p} OR a.hoofdgroep ILIKE ${p} OR array_to_string(a.vbn_codes, ' ') ILIKE ${p} OR a.growers::text ILIKE ${p})`);
  }
  values.push(Math.max(1, Math.min(1000, Number(limit) || 200)));
  const result = await dbQuery(
    `SELECT a.*, (SELECT count(*)::int FROM camp_items i WHERE i.article_id = a.id) AS campaigns
     FROM camp_articles a ${conditions.length ? `WHERE ${conditions.join(" AND ")}` : ""}
     ORDER BY a.updated_at DESC LIMIT $${values.length}`,
    values,
  );
  return result.rows.map(articleRow);
}

async function getArticle(id) {
  const result = await dbQuery("SELECT * FROM camp_articles WHERE id = $1", [id]);
  const article = articleRow(result.rows[0]);
  if (!article) return null;
  const history = await dbQuery(
    `SELECT i.id::int AS item_id, i.status, i.offer_price::float8 AS offer_price, i.planned_qty::float8 AS planned_qty, c.id::int AS campaign_id, c.offer_month, c.stage
     FROM camp_items i JOIN camp_campaigns c ON c.id = i.campaign_id WHERE i.article_id = $1 ORDER BY c.offer_month DESC`,
    [id],
  );
  const deals = await dbQuery(
    `SELECT d.*, c.offer_month, c.id AS campaign_id FROM camp_grower_deals d
     JOIN camp_items i ON i.id = d.item_id JOIN camp_campaigns c ON c.id = i.campaign_id
     WHERE i.article_id = $1 ORDER BY c.offer_month DESC, d.id`,
    [id],
  );
  return { ...article, history: history.rows, deals: deals.rows.map(dealRow) };
}

function articleFields(body) {
  const fields = {};
  if (body.name !== undefined) fields.name = text(body.name, 200);
  if (body.name_fr !== undefined) fields.name_fr = text(body.name_fr, 200);
  if (body.name_en !== undefined) fields.name_en = text(body.name_en, 200);
  if (body.vbn_codes !== undefined) fields.vbn_codes = [...new Set((Array.isArray(body.vbn_codes) ? body.vbn_codes : String(body.vbn_codes).split(/[\s,;]+/)).map((code) => text(code, 20)).filter(Boolean))].slice(0, 20);
  if (body.product_id !== undefined) fields.product_id = text(body.product_id, 60);
  if (body.hoofdgroep !== undefined) fields.hoofdgroep = text(body.hoofdgroep, 120);
  if (body.sales_group !== undefined) fields.sales_group = text(body.sales_group, 120);
  if (body.conditions !== undefined) fields.conditions = JSON.stringify(normalizeConditions(body.conditions));
  if (body.growers !== undefined) fields.growers = JSON.stringify(normalizeGrowers(body.growers));
  if (body.match_rules !== undefined) fields.match_rules = JSON.stringify(normalizeMatchRules(body.match_rules));
  if (body.notes !== undefined) fields.notes = text(body.notes, 4000);
  if (body.archived !== undefined) fields.archived = body.archived === true;
  return fields;
}

async function insertArticle(body, by) {
  const fields = articleFields(body);
  if (!fields.name) throw Object.assign(new Error("An article needs a name"), { status: 400 });
  // Sales matching starts from the VBN codes and growers it was made from.
  if (fields.match_rules === undefined) {
    const growers = body.growers ? normalizeGrowers(body.growers) : [];
    fields.match_rules = JSON.stringify(normalizeMatchRules({ vbn_codes: fields.vbn_codes || [], grower_codes: growers.map((grower) => grower.code).filter(Boolean) }));
  }
  const keys = Object.keys(fields);
  const values = keys.map((key) => fields[key]);
  values.push(by, by);
  const result = await dbQuery(
    `INSERT INTO camp_articles (${keys.join(", ")}, created_by, updated_by) VALUES (${keys.map((_, index) => `$${index + 1}`).join(", ")}, $${keys.length + 1}, $${keys.length + 2}) RETURNING *`,
    values,
  );
  return articleRow(result.rows[0]);
}

async function updateRow(table, id, fields, by) {
  const keys = Object.keys(fields);
  if (!keys.length) {
    const result = await dbQuery(`SELECT * FROM ${table} WHERE id = $1`, [id]);
    return result.rows[0] || null;
  }
  const values = keys.map((key) => fields[key]);
  values.push(by, id);
  const result = await dbQuery(
    `UPDATE ${table} SET ${keys.map((key, index) => `${key} = $${index + 1}`).join(", ")}, updated_by = $${keys.length + 1}, updated_at = now() WHERE id = $${keys.length + 2} RETURNING *`,
    values,
  );
  return result.rows[0] || null;
}

function dealFields(body) {
  const fields = {};
  if (body.grower_name !== undefined) fields.grower_name = text(body.grower_name, 160);
  if (body.grower_code !== undefined) fields.grower_code = text(body.grower_code, 40);
  if (body.grower_gln !== undefined) fields.grower_gln = text(body.grower_gln, 20);
  if (body.grower_fh !== undefined) fields.grower_fh = text(body.grower_fh, 20);
  if (body.status !== undefined && GROWER_DEAL_STATUSES.includes(body.status)) fields.status = body.status;
  if (body.agreed_price !== undefined) fields.agreed_price = num(body.agreed_price);
  if (body.qty_plan !== undefined) {
    fields.qty_plan = JSON.stringify((Array.isArray(body.qty_plan) ? body.qty_plan : []).slice(0, 60).map((row) => ({
      week: text(row?.week, 10),
      qty: num(row?.qty),
      note: text(row?.note, 200),
    })).filter((row) => row.week || row.qty !== null));
  }
  if (body.conditions !== undefined) fields.conditions = JSON.stringify(normalizeConditions(body.conditions));
  if (body.valid_from !== undefined) fields.valid_from = isoDay(body.valid_from);
  if (body.valid_to !== undefined) fields.valid_to = isoDay(body.valid_to);
  if (body.next_action_at !== undefined) fields.next_action_at = isoDay(body.next_action_at);
  if (body.next_action !== undefined) fields.next_action = text(body.next_action, 300);
  return fields;
}

async function dealContext(dealId) {
  const result = await dbQuery(
    `SELECT d.*, a.id AS article_id, a.name AS article_name, c.id AS campaign_id, c.offer_month
     FROM camp_grower_deals d JOIN camp_items i ON i.id = d.item_id JOIN camp_articles a ON a.id = i.article_id
     JOIN camp_campaigns c ON c.id = i.campaign_id WHERE d.id = $1`,
    [dealId],
  );
  return dealRow(result.rows[0]);
}

async function listDeals({ campaignId, grower, openOnly } = {}) {
  const values = [];
  const conditions = [];
  if (campaignId) {
    values.push(campaignId);
    conditions.push(`c.id = $${values.length}`);
  }
  if (grower) {
    values.push(`%${String(grower).replace(/[\\%_]/g, (match) => `\\${match}`)}%`);
    conditions.push(`(d.grower_name ILIKE $${values.length} OR d.grower_code ILIKE $${values.length})`);
  }
  if (openOnly) conditions.push("d.status NOT IN ('closed', 'declined')");
  const result = await dbQuery(
    `SELECT d.*, a.id AS article_id, a.name AS article_name, c.id AS campaign_id, c.offer_month,
       (SELECT count(*)::int FROM camp_grower_contacts k WHERE k.deal_id = d.id) AS contacts,
       (SELECT max(k.at) FROM camp_grower_contacts k WHERE k.deal_id = d.id) AS last_contact_at
     FROM camp_grower_deals d JOIN camp_items i ON i.id = d.item_id JOIN camp_articles a ON a.id = i.article_id
     JOIN camp_campaigns c ON c.id = i.campaign_id
     ${conditions.length ? `WHERE ${conditions.join(" AND ")}` : ""}
     ORDER BY c.offer_month DESC, d.grower_name, a.name LIMIT 2000`,
    values,
  );
  return result.rows.map(dealRow);
}

// ---------------------------------------------------------------------------
// HTTP

export function createCampaignRoutes({ sendJson, readRequestJson, requirePermission, hasUserPermission, PERMISSIONS, hoofdgroepFor, growerCodeFor }) {
  const canManage = (user) => hasUserPermission(user, PERMISSIONS.CAMPAIGN_MANAGE);
  const needManage = (res, user) => requirePermission(res, user, PERMISSIONS.CAMPAIGN_MANAGE);

  return async function handleCampaignRoute(req, res, url, requestUser) {
    const pathname = url.pathname;
    if (pathname !== "/api/campaigns" && !pathname.startsWith("/api/campaigns/")) return false;
    if (!requirePermission(res, requestUser, PERMISSIONS.CAMPAIGN_VIEW)) return true;
    if (!isDatabaseEnabled()) {
      sendJson(res, 503, { error: "Campagnes need the database" });
      return true;
    }
    const by = requestUser.username;
    const method = req.method;
    let match;
    try {
      // ---- Settings
      if (pathname === "/api/campaigns/settings") {
        if (method === "GET") {
          sendJson(res, 200, { settings: await readCampaignSettings(), can_manage: canManage(requestUser), stages: CAMPAIGN_STAGES,
            item_statuses: CAMPAIGN_ITEM_STATUSES, deal_statuses: GROWER_DEAL_STATUSES, contact_kinds: GROWER_CONTACT_KINDS, feature_labels: VBN_FEATURE_LABELS });
          return true;
        }
        if (method === "PUT") {
          if (!needManage(res, requestUser)) return true;
          sendJson(res, 200, { settings: await writeCampaignSettings(await readRequestJson(req), by) });
          return true;
        }
      }

      // ---- Sources: our invoices and product export (read only)
      if (pathname === "/api/campaigns/source/invoices" && method === "GET") {
        const rows = await searchInkoopArticles({
          product: url.searchParams.get("product"),
          grower: url.searchParams.get("grower"),
          vbn: url.searchParams.get("vbn"),
          months: url.searchParams.get("months"),
          limit: url.searchParams.get("limit"),
        });
        const growerCodes = growerCodeFor ? await growerCodeFor(rows) : new Map();
        const groupOf = hoofdgroepFor ? await hoofdgroepFor() : () => "";
        sendJson(res, 200, {
          rows: rows.map((row) => ({
            ...row,
            hoofdgroep: groupOf(row.description),
            grower_code: growerCodes.get(`${row.supplier_gln || ""}|${row.supplier_fh_number || ""}`) || "",
          })),
        });
        return true;
      }
      if (pathname === "/api/campaigns/source/products" && method === "GET") {
        sendJson(res, 200, { rows: await searchInkoopProductCards({ q: url.searchParams.get("q"), limit: url.searchParams.get("limit") }) });
        return true;
      }

      // ---- Master articles
      if (pathname === "/api/campaigns/articles") {
        if (method === "GET") {
          sendJson(res, 200, { articles: await listArticles({ q: url.searchParams.get("q"), includeArchived: url.searchParams.get("archived") === "1" }) });
          return true;
        }
        if (method === "POST") {
          if (!needManage(res, requestUser)) return true;
          const body = await readRequestJson(req);
          if (!body.hoofdgroep && hoofdgroepFor) body.hoofdgroep = (await hoofdgroepFor())(body.name);
          sendJson(res, 200, { article: await insertArticle(body, by) });
          return true;
        }
      }
      if ((match = pathname.match(/^\/api\/campaigns\/articles\/(\d+)$/))) {
        const id = Number(match[1]);
        if (method === "GET") {
          const article = await getArticle(id);
          sendJson(res, article ? 200 : 404, article ? { article } : { error: "Article not found" });
          return true;
        }
        if (method === "PATCH") {
          if (!needManage(res, requestUser)) return true;
          const row = await updateRow("camp_articles", id, articleFields(await readRequestJson(req)), by);
          sendJson(res, row ? 200 : 404, row ? { article: articleRow(row) } : { error: "Article not found" });
          return true;
        }
      }

      // ---- Campaigns
      if (pathname === "/api/campaigns") {
        if (method === "GET") {
          sendJson(res, 200, { campaigns: await listCampaigns() });
          return true;
        }
        if (method === "POST") {
          if (!needManage(res, requestUser)) return true;
          const body = await readRequestJson(req);
          const offerMonth = text(body.offer_month, 7);
          if (!isMonth(offerMonth)) {
            sendJson(res, 400, { error: "Offer month must be YYYY-MM" });
            return true;
          }
          const settings = await readCampaignSettings();
          const timeline = defaultCampaignTimeline(offerMonth, body.lead_months || settings.lead_months);
          const result = await dbQuery(
            `INSERT INTO camp_campaigns (offer_month, title_fr, title_en, campaign_start, campaign_end, deadlines, notes, created_by, updated_by)
             VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $8) ON CONFLICT (offer_month) DO NOTHING RETURNING *`,
            [offerMonth, text(body.title_fr, 200), text(body.title_en, 200), timeline.campaign_start, timeline.campaign_end,
              JSON.stringify(timeline.deadlines), text(body.notes, 4000), by],
          );
          if (!result.rows[0]) {
            sendJson(res, 409, { error: `There is already a campaign for ${offerMonth}` });
            return true;
          }
          sendJson(res, 200, { campaign: campaignRow(result.rows[0]) });
          return true;
        }
      }
      if ((match = pathname.match(/^\/api\/campaigns\/(\d+)$/))) {
        const id = Number(match[1]);
        if (method === "GET") {
          const campaign = await getCampaign(id);
          if (!campaign) {
            sendJson(res, 404, { error: "Campaign not found" });
            return true;
          }
          sendJson(res, 200, { campaign, items: await getCampaignItems(id) });
          return true;
        }
        if (method === "PATCH") {
          if (!needManage(res, requestUser)) return true;
          const body = await readRequestJson(req);
          const fields = {};
          if (body.title_fr !== undefined) fields.title_fr = text(body.title_fr, 200);
          if (body.title_en !== undefined) fields.title_en = text(body.title_en, 200);
          if (body.campaign_start !== undefined) fields.campaign_start = isoDay(body.campaign_start);
          if (body.campaign_end !== undefined) fields.campaign_end = isoDay(body.campaign_end);
          if (body.stage !== undefined && CAMPAIGN_STAGES.includes(body.stage)) fields.stage = body.stage;
          if (body.notes !== undefined) fields.notes = text(body.notes, 4000);
          if (body.deadlines !== undefined) {
            const deadlines = {};
            for (const [key, day] of Object.entries(body.deadlines || {}).slice(0, 12)) if (isoDay(day)) deadlines[text(key, 30)] = isoDay(day);
            fields.deadlines = JSON.stringify(deadlines);
          }
          const row = await updateRow("camp_campaigns", id, fields, by);
          sendJson(res, row ? 200 : 404, row ? { campaign: campaignRow(row) } : { error: "Campaign not found" });
          return true;
        }
        if (method === "DELETE") {
          if (!needManage(res, requestUser)) return true;
          // Only an empty campaign is removed; one with articles is set to
          // "cancelled" instead, so nothing that was planned is lost.
          const items = await dbQuery("SELECT count(*)::int AS n FROM camp_items WHERE campaign_id = $1", [id]);
          if (items.rows[0].n > 0) {
            sendJson(res, 409, { error: "This campaign has articles -- set its stage to Cancelled instead" });
            return true;
          }
          await dbQuery("DELETE FROM camp_campaigns WHERE id = $1", [id]);
          sendJson(res, 200, { ok: true });
          return true;
        }
      }
      // Copy the articles of an earlier campaign (shortlist again).
      if ((match = pathname.match(/^\/api\/campaigns\/(\d+)\/copy-from\/(\d+)$/)) && method === "POST") {
        if (!needManage(res, requestUser)) return true;
        const result = await dbQuery(
          `INSERT INTO camp_items (campaign_id, article_id, status, offer_price, planned_qty, sort_order, created_by, updated_by)
           SELECT $1, article_id, 'shortlist', offer_price, planned_qty, sort_order, $3, $3 FROM camp_items
           WHERE campaign_id = $2 AND status <> 'dropped'
           ON CONFLICT (campaign_id, article_id) DO NOTHING`,
          [Number(match[1]), Number(match[2]), by],
        );
        sendJson(res, 200, { added: result.rowCount });
        return true;
      }

      // ---- Items (an article in a campaign)
      if ((match = pathname.match(/^\/api\/campaigns\/(\d+)\/items$/)) && method === "POST") {
        if (!needManage(res, requestUser)) return true;
        const campaignId = Number(match[1]);
        const body = await readRequestJson(req);
        let articleId = Number(body.article_id) || 0;
        if (!articleId && body.article) {
          if (!body.article.hoofdgroep && hoofdgroepFor) body.article.hoofdgroep = (await hoofdgroepFor())(body.article.name);
          articleId = (await insertArticle(body.article, by)).id;
        }
        if (!articleId) {
          sendJson(res, 400, { error: "Choose or make an article" });
          return true;
        }
        const order = await dbQuery("SELECT COALESCE(max(sort_order), 0) + 1 AS n FROM camp_items WHERE campaign_id = $1", [campaignId]);
        const result = await dbQuery(
          `INSERT INTO camp_items (campaign_id, article_id, status, offer_price, planned_qty, sort_order, notes, created_by, updated_by)
           VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $8) ON CONFLICT (campaign_id, article_id) DO NOTHING RETURNING id`,
          [campaignId, articleId, CAMPAIGN_ITEM_STATUSES.includes(body.status) ? body.status : "shortlist", num(body.offer_price), num(body.planned_qty),
            order.rows[0].n, text(body.notes, 2000), by],
        );
        if (!result.rows[0]) {
          sendJson(res, 409, { error: "This article is already in the campaign" });
          return true;
        }
        sendJson(res, 200, { item_id: Number(result.rows[0].id), article_id: articleId });
        return true;
      }
      if ((match = pathname.match(/^\/api\/campaigns\/items\/(\d+)$/))) {
        const id = Number(match[1]);
        if (method === "PATCH") {
          if (!needManage(res, requestUser)) return true;
          const body = await readRequestJson(req);
          const fields = {};
          if (body.status !== undefined && CAMPAIGN_ITEM_STATUSES.includes(body.status)) fields.status = body.status;
          if (body.offer_price !== undefined) fields.offer_price = num(body.offer_price);
          if (body.planned_qty !== undefined) fields.planned_qty = num(body.planned_qty);
          if (body.sort_order !== undefined) fields.sort_order = Math.round(Number(body.sort_order) || 0);
          if (body.notes !== undefined) fields.notes = text(body.notes, 2000);
          const row = await updateRow("camp_items", id, fields, by);
          sendJson(res, row ? 200 : 404, row ? { ok: true } : { error: "Item not found" });
          return true;
        }
        if (method === "DELETE") {
          if (!needManage(res, requestUser)) return true;
          // With grower deals it is only marked dropped (deals stay).
          const deals = await dbQuery("SELECT count(*)::int AS n FROM camp_grower_deals WHERE item_id = $1", [id]);
          if (deals.rows[0].n > 0) {
            await updateRow("camp_items", id, { status: "dropped" }, by);
            sendJson(res, 200, { ok: true, dropped: true });
            return true;
          }
          await dbQuery("DELETE FROM camp_items WHERE id = $1", [id]);
          sendJson(res, 200, { ok: true });
          return true;
        }
      }

      // ---- Grower deals
      if (pathname === "/api/campaigns/deals" && method === "GET") {
        sendJson(res, 200, {
          deals: await listDeals({
            campaignId: Number(url.searchParams.get("campaign_id")) || null,
            grower: url.searchParams.get("grower"),
            openOnly: url.searchParams.get("open") === "1",
          }),
        });
        return true;
      }
      if ((match = pathname.match(/^\/api\/campaigns\/items\/(\d+)\/deals$/)) && method === "POST") {
        if (!needManage(res, requestUser)) return true;
        const fields = dealFields(await readRequestJson(req));
        if (!fields.grower_name && !fields.grower_code) {
          sendJson(res, 400, { error: "A deal needs a grower" });
          return true;
        }
        fields.item_id = Number(match[1]);
        const keys = Object.keys(fields);
        const values = keys.map((key) => fields[key]);
        values.push(by);
        const result = await dbQuery(
          `INSERT INTO camp_grower_deals (${keys.join(", ")}, created_by, updated_by) VALUES (${keys.map((_, index) => `$${index + 1}`).join(", ")}, $${keys.length + 1}, $${keys.length + 1}) RETURNING id`,
          values,
        );
        const deal = await dealContext(Number(result.rows[0].id));
        await taskForDealNextAction(deal, deal.article_name, deal.offer_month, by);
        sendJson(res, 200, { deal });
        return true;
      }
      if ((match = pathname.match(/^\/api\/campaigns\/deals\/(\d+)$/))) {
        const id = Number(match[1]);
        if (method === "GET") {
          const deal = await dealContext(id);
          if (!deal) {
            sendJson(res, 404, { error: "Deal not found" });
            return true;
          }
          const contacts = await dbQuery("SELECT id, at, kind, note, by_user FROM camp_grower_contacts WHERE deal_id = $1 ORDER BY at DESC", [id]);
          sendJson(res, 200, { deal, contacts: contacts.rows });
          return true;
        }
        if (method === "PATCH") {
          if (!needManage(res, requestUser)) return true;
          const before = await dealContext(id);
          if (!before) {
            sendJson(res, 404, { error: "Deal not found" });
            return true;
          }
          const fields = dealFields(await readRequestJson(req));
          await updateRow("camp_grower_deals", id, fields, by);
          const deal = await dealContext(id);
          // The status change goes into the contact log too.
          if (fields.status && fields.status !== before.status) {
            await dbQuery("INSERT INTO camp_grower_contacts (deal_id, kind, note, by_user) VALUES ($1, 'note', $2, $3)", [id, `Status: ${before.status} -> ${fields.status}`, by]);
          }
          if (deal.next_action_at && deal.next_action_at !== before.next_action_at) await taskForDealNextAction(deal, deal.article_name, deal.offer_month, by);
          sendJson(res, 200, { deal });
          return true;
        }
      }
      if ((match = pathname.match(/^\/api\/campaigns\/deals\/(\d+)\/contacts$/)) && method === "POST") {
        if (!needManage(res, requestUser)) return true;
        const body = await readRequestJson(req);
        const note = text(body.note, 4000);
        if (!note) {
          sendJson(res, 400, { error: "Write what was said or done" });
          return true;
        }
        await dbQuery(
          "INSERT INTO camp_grower_contacts (deal_id, at, kind, note, by_user) VALUES ($1, COALESCE($2::timestamptz, now()), $3, $4, $5)",
          [Number(match[1]), body.at || null, GROWER_CONTACT_KINDS.includes(body.kind) ? body.kind : "note", note, by],
        );
        sendJson(res, 200, { ok: true });
        return true;
      }

      sendJson(res, 404, { error: "Unknown campaigns route" });
      return true;
    } catch (error) {
      console.error("Campaigns route failed:", pathname, error?.message || error);
      sendJson(res, error?.status || 500, { error: error?.message || "Campaigns error" });
      return true;
    }
  };
}
