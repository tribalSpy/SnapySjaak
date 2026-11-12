// Campagnes media: the photo library (a Social Drive folder), themes,
// renders (photos / videos made on the BEAST PC) and posts.
//
// Render (this server) only keeps the records and queues jobs; every heavy
// step -- reading photos from Drive, thumbnails, ffmpeg, AI, captions,
// uploading the results to Drive -- runs on the BEAST PC's media-poller-app
// (job types media_*; db.js keeps them away from the other pollers). No
// image or video bytes pass through here: previews and downloads are Drive
// links.
//
// Posts can be recorded by hand (posted from the phone / browser, link
// pasted back) until the platforms' APIs are approved; stats too.
import { createLlmJob, dbQuery, isDatabaseEnabled } from "./db.js";

// Everything the BEAST knows how to make. engine "ai" = phase 4 (ComfyUI).
export const MEDIA_RECIPES = [
  { key: "zoom_in", label: "Slow zoom in", kind: "video", engine: "ffmpeg", photos: "1" },
  { key: "zoom_out", label: "Slow zoom out", kind: "video", engine: "ffmpeg", photos: "1" },
  { key: "pan", label: "Pan across", kind: "video", engine: "ffmpeg", photos: "1" },
  { key: "detail_tour", label: "Detail tour (zoom into details)", kind: "video", engine: "ffmpeg", photos: "1" },
  { key: "slideshow", label: "Slideshow", kind: "video", engine: "ffmpeg", photos: "many" },
  { key: "offer_reel", label: "Offer reel (articles with name cards)", kind: "video", engine: "ffmpeg", photos: "many" },
  { key: "photo_post", label: "Photo post (framed, with text)", kind: "photo", engine: "ffmpeg", photos: "1" },
  { key: "carousel", label: "Carousel (one card per photo)", kind: "photo", engine: "ffmpeg", photos: "many" },
  { key: "catalogue_card", label: "Catalogue card (article + conditions)", kind: "photo", engine: "ffmpeg", photos: "1" },
  { key: "ai_flower_drop", label: "AI: flowers dropping, slow motion", kind: "video", engine: "ai", photos: "1" },
  { key: "ai_living", label: "AI: living photo (gentle motion)", kind: "video", engine: "ai", photos: "1" },
  { key: "ai_art", label: "AI: photo → art drawing", kind: "video", engine: "ai", photos: "1" },
];
const RECIPE_KEYS = new Set(MEDIA_RECIPES.map((recipe) => recipe.key));

// AI recipes can be queued once a BEAST has said (in the last 14 days) that
// it has the AI part switched on (capability media_ai).
async function aiAvailable() {
  try {
    const result = await dbQuery("SELECT 1 FROM llm_agents WHERE capabilities ? 'media_ai' AND last_seen_at > now() - interval '14 days' LIMIT 1");
    return Boolean(result.rows[0]);
  } catch {
    return false;
  }
}

export const MEDIA_PLATFORMS = ["instagram", "facebook", "tiktok", "youtube", "linkedin", "pinterest"];
const POST_STATUSES = ["planned", "approved", "published", "failed", "cancelled"];
const VARIANTS = ["9x16", "4x5", "1x1"];

const DEFAULT_THEMES = [
  { key: "flower_drop", name: "Flower Drop", recipes: ["ai_flower_drop", "zoom_in", "slideshow"],
    caption_prompt: "Slow-motion falling flowers: dreamy, short, poetic. Mention the flower by name.", hashtags: "#flowers #slowmotion #fleurs #bloemen" },
  { key: "plant_of_month", name: "Plant of the Month", recipes: ["detail_tour", "carousel", "photo_post"],
    caption_prompt: "Plant of the month: 2-3 small surprising details about this plant (leaves, origin, care), friendly and informative.", hashtags: "#plantofthemonth #plantedumois #plants #plantes" },
  { key: "deco_blooms", name: "Deco & Blooms", recipes: ["slideshow", "carousel", "photo_post"],
    caption_prompt: "Decoration ideas with these flowers: how florists and shops can use them in arrangements.", hashtags: "#flowerdeco #decoration #fleuriste #florist" },
  { key: "world_flowers", name: "World Flowers", recipes: ["zoom_out", "pan", "photo_post"],
    caption_prompt: "Where in the world this flower comes from and one fact about its origin.", hashtags: "#worldflowers #fleursdumonde #flowers" },
  { key: "strange_unusual", name: "Strange & Unusual", recipes: ["zoom_in", "detail_tour", "ai_living"],
    caption_prompt: "Did you know / Le saviez-vous: one strange or unusual fact about this flower or plant.", hashtags: "#unusualflowers #lesaviezvous #didyouknow" },
  { key: "photo_to_art", name: "Photo → Art", recipes: ["ai_art", "photo_post"],
    caption_prompt: "From photo to art: the flower turned into a drawing. Short and artistic.", hashtags: "#flowerart #botanicalart #artfloral" },
  { key: "seasonal", name: "Seasonal", recipes: ["slideshow", "photo_post"],
    caption_prompt: "Seasonal moment (Valentine, Mother's Day, Christmas...): warm and inviting, linked to the season.", hashtags: "#seasonalflowers #fleursdesaison" },
  { key: "offer_month", name: "Offer of the Month", recipes: ["offer_reel", "carousel", "catalogue_card"],
    caption_prompt: "Our offer for next month: list the highlights, invite customers to order. No prices unless given.", hashtags: "#offer #offre #wholesaleflowers #grossistefleurs" },
];

function text(value, max = 500) {
  return String(value ?? "").trim().slice(0, max);
}

function num(value) {
  if (value === null || value === undefined || value === "") return null;
  const parsed = Number(String(value).replace(",", "."));
  return Number.isFinite(parsed) ? parsed : null;
}

function idList(value) {
  return [...new Set((Array.isArray(value) ? value : []).map((item) => Number(item)).filter((item) => Number.isInteger(item) && item > 0))].slice(0, 60);
}

function textList(value, max = 60, count = 40) {
  return [...new Set((Array.isArray(value) ? value : String(value || "").split(/\s*,\s*/)).map((item) => text(item, max)).filter(Boolean))].slice(0, count);
}

// ---------------------------------------------------------------------------
// Settings (camp_settings, key "media")

function normalizeMediaSettings(value) {
  const source = value && typeof value === "object" ? value : {};
  const folder = (id) => text(id, 120).replace(/^.*\/folders\//, "").replace(/[?#].*$/, "");
  return {
    // The Social Drive folder people drop HD photos in (subfolders = tags).
    library_folder_id: folder(source.library_folder_id),
    // Where the BEAST puts what it made (and the library thumbnails).
    output_folder_id: folder(source.output_folder_id),
    // Outputs / thumbnails readable by link, so the app can preview them
    // without passing the files through Render.
    share_outputs: source.share_outputs !== false,
    brand_name: text(source.brand_name || "Sjaak van der Vijver", 80),
    accent: /^#[0-9a-f]{6}$/i.test(source.accent || "") ? source.accent : "#1f7a5a",
    default_variants: textList(source.default_variants?.length ? source.default_variants : ["9x16"], 8, 3).filter((variant) => VARIANTS.includes(variant)),
  };
}

export async function readMediaSettings() {
  if (!isDatabaseEnabled()) return normalizeMediaSettings(null);
  const result = await dbQuery("SELECT value FROM camp_settings WHERE key = 'media'");
  return normalizeMediaSettings(result.rows[0]?.value);
}

async function writeMediaSettings(value, by) {
  const normalized = normalizeMediaSettings(value);
  await dbQuery(
    `INSERT INTO camp_settings (key, value, updated_by, updated_at) VALUES ('media', $1, $2, now())
     ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_by = EXCLUDED.updated_by, updated_at = now()`,
    [JSON.stringify(normalized), by || ""],
  );
  return normalized;
}

// ---------------------------------------------------------------------------
// Themes

let themesSeeded = false;
async function seedThemes() {
  if (themesSeeded) return;
  for (const [index, theme] of DEFAULT_THEMES.entries()) {
    await dbQuery(
      `INSERT INTO media_themes (key, name, recipes, caption_prompt, hashtags, sort_order)
       VALUES ($1, $2, $3, $4, $5, $6) ON CONFLICT (key) DO NOTHING`,
      [theme.key, theme.name, theme.recipes, theme.caption_prompt, theme.hashtags, index + 1],
    );
  }
  themesSeeded = true;
}

async function listThemes() {
  await seedThemes();
  const result = await dbQuery("SELECT * FROM media_themes ORDER BY sort_order, name");
  return result.rows.map(themeRow);
}

function themeRow(row) {
  return {
    id: Number(row.id),
    key: row.key,
    name: row.name,
    description: row.description,
    recipes: row.recipes || [],
    caption_prompt: row.caption_prompt,
    hashtags: row.hashtags,
    look: row.look || {},
    music: row.music,
    cadence: row.cadence,
    autopilot: row.autopilot,
    active: row.active,
    sort_order: row.sort_order,
  };
}

// ---------------------------------------------------------------------------
// Rows

function assetRow(row) {
  return {
    id: Number(row.id),
    drive_id: row.drive_id,
    name: row.name,
    folder_path: row.folder_path,
    mime: row.mime,
    width: row.width,
    height: row.height,
    size: row.size === null ? null : Number(row.size),
    taken_at: row.taken_at,
    thumb_id: row.thumb_id,
    tags: row.tags || [],
    article_ids: (row.article_ids || []).map(Number),
    theme_keys: row.theme_keys || [],
    missing: row.missing,
    seen_at: row.seen_at,
  };
}

function itemRow(row) {
  return {
    id: Number(row.id),
    title: row.title,
    theme_key: row.theme_key,
    recipe: row.recipe,
    article_ids: (row.article_ids || []).map(Number),
    campaign_id: row.campaign_id === null ? null : Number(row.campaign_id),
    asset_ids: (row.asset_ids || []).map(Number),
    options: row.options || {},
    status: row.status,
    job_id: row.job_id,
    outputs: row.outputs || [],
    caption_fr: row.caption_fr,
    caption_en: row.caption_en,
    hashtags: row.hashtags,
    error: row.error,
    render_seconds: row.render_seconds === null ? null : Number(row.render_seconds),
    created_by: row.created_by,
    created_at: row.created_at,
    updated_at: row.updated_at,
    posts: row.posts === undefined ? undefined : row.posts || [],
  };
}

function postRow(row) {
  return {
    id: Number(row.id),
    media_item_id: row.media_item_id === null ? null : Number(row.media_item_id),
    platform: row.platform,
    scheduled_at: row.scheduled_at,
    posted_at: row.posted_at,
    status: row.status,
    mode: row.mode,
    post_url: row.post_url,
    caption: row.caption,
    approved_by: row.approved_by,
    notes: row.notes,
    created_by: row.created_by,
    updated_at: row.updated_at,
    item_title: row.item_title,
    theme_key: row.theme_key,
    recipe: row.recipe,
    outputs: row.outputs,
    stats: row.stats || [],
  };
}

// ---------------------------------------------------------------------------
// Jobs for the BEAST

async function queueScan(by) {
  const settings = await readMediaSettings();
  if (!settings.library_folder_id) throw Object.assign(new Error("Set the Social library Drive folder first (Settings)"), { status: 400 });
  if (!settings.output_folder_id) throw Object.assign(new Error("Set the output Drive folder first (Settings)"), { status: 400 });
  const pending = await dbQuery("SELECT id FROM llm_jobs WHERE job_type = 'media_scan' AND status IN ('pending', 'claimed') LIMIT 1");
  if (pending.rows[0]) return { job_id: pending.rows[0].id, already: true };
  const job = await createLlmJob({
    job_type: "media_scan",
    created_by: by,
    priority: 20,
    max_attempts: 2,
    payload_json: { library_folder_id: settings.library_folder_id, output_folder_id: settings.output_folder_id, share_outputs: settings.share_outputs },
  });
  return { job_id: job.id };
}

// Everything the BEAST needs to make one item: the photos (Drive ids), the
// recipe and options, the article texts (FR / EN / conditions) and the
// theme's caption prompt.
async function queueRender(itemId, by) {
  const itemResult = await dbQuery("SELECT * FROM media_items WHERE id = $1", [itemId]);
  const item = itemRow(itemResult.rows[0] || {});
  if (!itemResult.rows[0]) throw Object.assign(new Error("Item not found"), { status: 404 });
  const recipe = MEDIA_RECIPES.find((entry) => entry.key === item.recipe);
  if (recipe?.engine === "ai" && !(await aiAvailable())) throw Object.assign(new Error("The BEAST has no AI part yet (setup_media_pc.bat, AI step; config ai.enabled)"), { status: 400 });
  const settings = await readMediaSettings();
  if (!settings.output_folder_id) throw Object.assign(new Error("Set the output Drive folder first (Settings)"), { status: 400 });
  const assets = item.asset_ids.length
    ? (await dbQuery("SELECT id, drive_id, name, width, height FROM media_assets WHERE id = ANY($1::bigint[])", [item.asset_ids])).rows
    : [];
  const ordered = item.asset_ids.map((id) => assets.find((asset) => Number(asset.id) === id)).filter(Boolean);
  if (!ordered.length) throw Object.assign(new Error("Choose at least one photo"), { status: 400 });
  const articles = item.article_ids.length
    ? (await dbQuery("SELECT id, name, name_fr, name_en, hoofdgroep, conditions FROM camp_articles WHERE id = ANY($1::bigint[])", [item.article_ids])).rows
    : [];
  let offerPrices = {};
  if (item.campaign_id && item.options.show_price) {
    const prices = await dbQuery("SELECT article_id, offer_price::float8 AS offer_price FROM camp_items WHERE campaign_id = $1", [item.campaign_id]);
    offerPrices = Object.fromEntries(prices.rows.map((row) => [Number(row.article_id), row.offer_price]));
  }
  const themeResult = await dbQuery("SELECT * FROM media_themes WHERE key = $1", [item.theme_key]);
  const theme = themeResult.rows[0] ? themeRow(themeResult.rows[0]) : null;
  const job = await createLlmJob({
    // AI waits for the BEAST's night window unless asked for now.
    job_type: recipe?.engine === "ai" ? (item.options.now ? "media_ai_now" : "media_ai") : "media_render",
    created_by: by,
    priority: 10,
    max_attempts: 2,
    payload_json: {
      item_id: item.id,
      recipe: item.recipe,
      kind: recipe?.kind || "video",
      options: { variants: settings.default_variants, ...item.options },
      assets: ordered.map((asset) => ({ id: Number(asset.id), drive_id: asset.drive_id, name: asset.name, width: asset.width, height: asset.height })),
      articles: item.article_ids.map((id) => articles.find((article) => Number(article.id) === id)).filter(Boolean).map((article) => ({
        id: Number(article.id),
        name: article.name,
        name_fr: article.name_fr,
        name_en: article.name_en,
        hoofdgroep: article.hoofdgroep,
        conditions: article.conditions || {},
        offer_price: offerPrices[Number(article.id)] ?? null,
      })),
      theme: theme ? { key: theme.key, name: theme.name, caption_prompt: theme.caption_prompt, hashtags: theme.hashtags, look: theme.look, music: theme.music } : null,
      brand: { name: settings.brand_name, accent: settings.accent },
      title: item.title,
      captions: { fr: item.caption_fr, en: item.caption_en, keep: Boolean(item.caption_fr || item.caption_en) },
      output: { folder_id: settings.output_folder_id, share: settings.share_outputs, subfolder: [theme?.name || item.theme_key || "Other", new Date().toISOString().slice(0, 7)] },
    },
  });
  await dbQuery("UPDATE media_items SET status = 'queued', job_id = $2, error = '', updated_at = now() WHERE id = $1", [item.id, job.id]);
  return { job_id: job.id };
}

// /api/llm/jobs/:id/result for a media_* job (called from index.js).
export async function handleMediaJobResult(job) {
  try {
    const result = job.result_json || {};
    if (job.job_type === "media_scan") return;
    const itemId = Number(job.payload_json?.item_id);
    if (!itemId) return;
    const outputs = (Array.isArray(result.outputs) ? result.outputs : []).slice(0, 60).map((output) => ({
      variant: text(output.variant, 20),
      kind: text(output.kind, 10),
      drive_id: text(output.drive_id, 120),
      name: text(output.name, 200),
      mime: text(output.mime, 60),
      width: num(output.width),
      height: num(output.height),
      duration: num(output.duration),
      size: num(output.size),
      index: num(output.index),
    })).filter((output) => output.drive_id);
    await dbQuery(
      `UPDATE media_items SET status = 'ready', outputs = $2, render_seconds = $3, error = '',
         caption_fr = CASE WHEN caption_fr = '' THEN $4 ELSE caption_fr END,
         caption_en = CASE WHEN caption_en = '' THEN $5 ELSE caption_en END,
         hashtags = CASE WHEN hashtags = '' THEN $6 ELSE hashtags END,
         updated_at = now()
       WHERE id = $1 AND job_id = $7`,
      [itemId, JSON.stringify(outputs), num(result.render_seconds), text(result.caption_fr, 4000), text(result.caption_en, 4000), text(result.hashtags, 1000), job.id],
    );
  } catch (error) {
    console.error("Media job result not saved:", job?.id, error?.message || error);
  }
}

export async function handleMediaJobFail(job, errorText) {
  try {
    const itemId = Number(job.payload_json?.item_id);
    if (!itemId) return;
    const final = job.status === "failed";
    await dbQuery("UPDATE media_items SET status = $2, error = $3, updated_at = now() WHERE id = $1 AND job_id = $4",
      [itemId, final ? "failed" : "queued", text(errorText || job.error_text, 2000), job.id]);
  } catch (error) {
    console.error("Media job failure not saved:", job?.id, error?.message || error);
  }
}

// The BEAST's library scan: upserts the photos it found; with complete=true
// every photo of the folder not in this scan is marked missing (kept, not
// deleted -- an item may still point at it).
async function saveScannedAssets(body) {
  const assets = (Array.isArray(body.assets) ? body.assets : []).slice(0, 2000);
  let saved = 0;
  for (const asset of assets) {
    const driveId = text(asset.drive_id, 120);
    if (!driveId) continue;
    const tags = textList(asset.tags, 60, 20);
    await dbQuery(
      `INSERT INTO media_assets (drive_id, name, folder_path, mime, width, height, size, taken_at, thumb_id, tags, scan_id, missing, seen_at)
       VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, false, now())
       ON CONFLICT (drive_id) DO UPDATE SET name = EXCLUDED.name, folder_path = EXCLUDED.folder_path, mime = EXCLUDED.mime,
         width = EXCLUDED.width, height = EXCLUDED.height, size = EXCLUDED.size, taken_at = COALESCE(EXCLUDED.taken_at, media_assets.taken_at),
         thumb_id = CASE WHEN EXCLUDED.thumb_id <> '' THEN EXCLUDED.thumb_id ELSE media_assets.thumb_id END,
         tags = (SELECT array_agg(DISTINCT tag) FROM unnest(media_assets.tags || EXCLUDED.tags) AS tag),
         scan_id = EXCLUDED.scan_id, missing = false, seen_at = now()`,
      [driveId, text(asset.name, 300), text(asset.folder_path, 500), text(asset.mime, 60), num(asset.width), num(asset.height), num(asset.size),
        asset.taken_at ? text(asset.taken_at, 40) : null, text(asset.thumb_id, 120), tags, text(body.scan_id, 80)],
    );
    saved += 1;
  }
  let missing = 0;
  if (body.complete === true && text(body.scan_id, 80)) {
    const result = await dbQuery("UPDATE media_assets SET missing = true WHERE scan_id IS DISTINCT FROM $1 AND NOT missing", [text(body.scan_id, 80)]);
    missing = result.rowCount;
  }
  return { saved, missing };
}

// The BEAST PC(s) running media-poller-app, and the media queue.
async function beastStatus() {
  const agents = await dbQuery(
    `SELECT agent_name, pc_name, status, version, capabilities, meta, last_seen_at,
       (last_seen_at > now() - interval '3 minutes') AS online
     FROM llm_agents WHERE capabilities ? 'media_render' ORDER BY last_seen_at DESC`,
  );
  const queue = await dbQuery(
    "SELECT job_type, status, count(*)::int AS n FROM llm_jobs WHERE job_type LIKE 'media\\_%' AND status IN ('pending', 'claimed') GROUP BY job_type, status",
  );
  const lastScan = await dbQuery("SELECT status, finished_at, created_at, error_text FROM llm_jobs WHERE job_type = 'media_scan' ORDER BY created_at DESC LIMIT 1");
  return { agents: agents.rows, queue: queue.rows, last_scan: lastScan.rows[0] || null };
}

// ---------------------------------------------------------------------------
// HTTP: /api/media/... (logged-in users; campaign permissions)

export function createMediaRoutes({ sendJson, readRequestJson, requirePermission, hasUserPermission, PERMISSIONS }) {
  const needManage = (res, user) => requirePermission(res, user, PERMISSIONS.CAMPAIGN_MANAGE);

  return async function handleMediaRoute(req, res, url, requestUser) {
    const pathname = url.pathname;
    if (!pathname.startsWith("/api/media/")) return false;
    if (!requirePermission(res, requestUser, PERMISSIONS.CAMPAIGN_VIEW)) return true;
    if (!isDatabaseEnabled()) {
      sendJson(res, 503, { error: "Media needs the database" });
      return true;
    }
    const by = requestUser.username;
    const method = req.method;
    let match;
    try {
      if (pathname === "/api/media/meta" && method === "GET") {
        sendJson(res, 200, {
          settings: await readMediaSettings(),
          recipes: await (async () => {
            const ai = await aiAvailable();
            return MEDIA_RECIPES.map((recipe) => ({ ...recipe, available: recipe.engine !== "ai" || ai }));
          })(),
          platforms: MEDIA_PLATFORMS,
          post_statuses: POST_STATUSES,
          variants: VARIANTS,
          can_manage: hasUserPermission(requestUser, PERMISSIONS.CAMPAIGN_MANAGE),
        });
        return true;
      }
      if (pathname === "/api/media/settings" && method === "PUT") {
        if (!needManage(res, requestUser)) return true;
        sendJson(res, 200, { settings: await writeMediaSettings(await readRequestJson(req), by) });
        return true;
      }
      if (pathname === "/api/media/beast" && method === "GET") {
        sendJson(res, 200, await beastStatus());
        return true;
      }

      // ---- Themes
      if (pathname === "/api/media/themes" && method === "GET") {
        sendJson(res, 200, { themes: await listThemes() });
        return true;
      }
      if (pathname === "/api/media/themes" && method === "POST") {
        if (!needManage(res, requestUser)) return true;
        const body = await readRequestJson(req);
        const key = text(body.key || body.name, 40).toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
        if (!key) {
          sendJson(res, 400, { error: "A theme needs a name" });
          return true;
        }
        const result = await dbQuery(
          `INSERT INTO media_themes (key, name, recipes, caption_prompt, hashtags, sort_order)
           VALUES ($1, $2, $3, $4, $5, (SELECT COALESCE(max(sort_order), 0) + 1 FROM media_themes)) ON CONFLICT (key) DO NOTHING RETURNING *`,
          [key, text(body.name, 80) || key, textList(body.recipes, 30, 12).filter((recipe) => RECIPE_KEYS.has(recipe)), text(body.caption_prompt, 2000), text(body.hashtags, 500)],
        );
        sendJson(res, result.rows[0] ? 200 : 409, result.rows[0] ? { theme: themeRow(result.rows[0]) } : { error: "There is already a theme with that name" });
        return true;
      }
      if ((match = pathname.match(/^\/api\/media\/themes\/(\d+)$/)) && method === "PATCH") {
        if (!needManage(res, requestUser)) return true;
        const body = await readRequestJson(req);
        const fields = {};
        if (body.name !== undefined) fields.name = text(body.name, 80);
        if (body.description !== undefined) fields.description = text(body.description, 1000);
        if (body.recipes !== undefined) fields.recipes = textList(body.recipes, 30, 12).filter((recipe) => RECIPE_KEYS.has(recipe));
        if (body.caption_prompt !== undefined) fields.caption_prompt = text(body.caption_prompt, 2000);
        if (body.hashtags !== undefined) fields.hashtags = text(body.hashtags, 500);
        if (body.music !== undefined) fields.music = text(body.music, 200);
        if (body.cadence !== undefined) fields.cadence = text(body.cadence, 120);
        if (body.active !== undefined) fields.active = body.active !== false;
        // Autopilot stays off until it is switched on here by hand.
        if (body.autopilot !== undefined) fields.autopilot = body.autopilot === true;
        const keys = Object.keys(fields);
        if (!keys.length) {
          sendJson(res, 400, { error: "Nothing to change" });
          return true;
        }
        const result = await dbQuery(
          `UPDATE media_themes SET ${keys.map((key, index) => `${key} = $${index + 2}`).join(", ")}, updated_at = now() WHERE id = $1 RETURNING *`,
          [Number(match[1]), ...keys.map((key) => fields[key])],
        );
        sendJson(res, result.rows[0] ? 200 : 404, result.rows[0] ? { theme: themeRow(result.rows[0]) } : { error: "Theme not found" });
        return true;
      }

      // ---- Library
      if (pathname === "/api/media/assets" && method === "GET") {
        const values = [];
        const conditions = [url.searchParams.get("missing") === "1" ? "true" : "NOT missing"];
        for (const word of String(url.searchParams.get("q") || "").split(/\s+/).filter(Boolean).slice(0, 6)) {
          values.push(`%${word.replace(/[\\%_]/g, (m) => `\\${m}`)}%`);
          conditions.push(`(name ILIKE $${values.length} OR folder_path ILIKE $${values.length} OR array_to_string(tags, ' ') ILIKE $${values.length})`);
        }
        const articleId = Number(url.searchParams.get("article_id"));
        if (articleId) {
          values.push(articleId);
          conditions.push(`$${values.length} = ANY(article_ids)`);
        }
        const theme = text(url.searchParams.get("theme"), 40);
        if (theme) {
          values.push(theme);
          conditions.push(`$${values.length} = ANY(theme_keys)`);
        }
        const result = await dbQuery(
          `SELECT * FROM media_assets WHERE ${conditions.join(" AND ")} ORDER BY COALESCE(taken_at, seen_at) DESC, id DESC LIMIT 600`,
          values,
        );
        sendJson(res, 200, { assets: result.rows.map(assetRow) });
        return true;
      }
      if (pathname === "/api/media/assets/scan" && method === "POST") {
        if (!needManage(res, requestUser)) return true;
        sendJson(res, 200, await queueScan(by));
        return true;
      }
      if ((match = pathname.match(/^\/api\/media\/assets\/(\d+)$/)) && method === "PATCH") {
        if (!needManage(res, requestUser)) return true;
        const body = await readRequestJson(req);
        const fields = {};
        if (body.tags !== undefined) fields.tags = textList(body.tags, 60, 30);
        if (body.article_ids !== undefined) fields.article_ids = idList(body.article_ids);
        if (body.theme_keys !== undefined) fields.theme_keys = textList(body.theme_keys, 40, 12);
        const keys = Object.keys(fields);
        if (!keys.length) {
          sendJson(res, 400, { error: "Nothing to change" });
          return true;
        }
        const result = await dbQuery(
          `UPDATE media_assets SET ${keys.map((key, index) => `${key} = $${index + 2}`).join(", ")} WHERE id = $1 RETURNING *`,
          [Number(match[1]), ...keys.map((key) => fields[key])],
        );
        sendJson(res, result.rows[0] ? 200 : 404, result.rows[0] ? { asset: assetRow(result.rows[0]) } : { error: "Photo not found" });
        return true;
      }

      // ---- Items (things the BEAST makes)
      if (pathname === "/api/media/items" && method === "GET") {
        const values = [];
        const conditions = [];
        const campaignId = Number(url.searchParams.get("campaign_id"));
        if (campaignId) {
          values.push(campaignId);
          conditions.push(`i.campaign_id = $${values.length}`);
        }
        const articleId = Number(url.searchParams.get("article_id"));
        if (articleId) {
          values.push(articleId);
          conditions.push(`$${values.length} = ANY(i.article_ids)`);
        }
        const result = await dbQuery(
          `SELECT i.*, COALESCE((SELECT jsonb_agg(jsonb_build_object('id', p.id, 'platform', p.platform, 'status', p.status, 'scheduled_at', p.scheduled_at, 'posted_at', p.posted_at, 'post_url', p.post_url) ORDER BY p.id)
             FROM social_posts p WHERE p.media_item_id = i.id), '[]'::jsonb) AS posts
           FROM media_items i ${conditions.length ? `WHERE ${conditions.join(" AND ")}` : ""}
           ORDER BY i.created_at DESC LIMIT 300`,
          values,
        );
        sendJson(res, 200, { items: result.rows.map(itemRow) });
        return true;
      }
      if (pathname === "/api/media/items" && method === "POST") {
        if (!needManage(res, requestUser)) return true;
        const body = await readRequestJson(req);
        if (!RECIPE_KEYS.has(body.recipe)) {
          sendJson(res, 400, { error: "Choose what to make" });
          return true;
        }
        if (MEDIA_RECIPES.find((recipe) => recipe.key === body.recipe)?.engine === "ai" && !(await aiAvailable())) {
          sendJson(res, 400, { error: "The BEAST has no AI part yet (setup_media_pc.bat, AI step; config ai.enabled)" });
          return true;
        }
        const options = body.options && typeof body.options === "object" ? body.options : {};
        const cleanOptions = {
          duration: Math.max(3, Math.min(60, num(options.duration) || 8)),
          variants: textList(options.variants, 8, 3).filter((variant) => VARIANTS.includes(variant)),
          text_fr: text(options.text_fr, 200),
          text_en: text(options.text_en, 200),
          show_price: options.show_price === true,
          show_conditions: options.show_conditions !== false,
          music: text(options.music, 200),
          now: options.now === true,
          ai_prompt: text(options.ai_prompt, 300),
          art_styles: textList(options.art_styles, 20, 3).filter((style) => ["sketch", "watercolor", "oil"].includes(style)),
          points: (Array.isArray(options.points) ? options.points : []).slice(0, 8)
            .map((point) => ({ x: Math.max(0, Math.min(1, num(point?.x) ?? 0.5)), y: Math.max(0, Math.min(1, num(point?.y) ?? 0.5)) })),
        };
        if (!cleanOptions.variants.length) delete cleanOptions.variants;
        const result = await dbQuery(
          `INSERT INTO media_items (title, theme_key, recipe, article_ids, campaign_id, asset_ids, options, caption_fr, caption_en, hashtags, created_by)
           VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11) RETURNING id`,
          [text(body.title, 200), text(body.theme_key, 40), body.recipe, idList(body.article_ids), Number(body.campaign_id) || null, idList(body.asset_ids),
            JSON.stringify(cleanOptions), text(body.caption_fr, 4000), text(body.caption_en, 4000), text(body.hashtags, 1000), by],
        );
        const itemId = Number(result.rows[0].id);
        try {
          const queued = await queueRender(itemId, by);
          sendJson(res, 200, { item_id: itemId, ...queued });
        } catch (error) {
          await dbQuery("UPDATE media_items SET status = 'failed', error = $2 WHERE id = $1", [itemId, text(error.message, 500)]);
          sendJson(res, error.status || 500, { error: error.message, item_id: itemId });
        }
        return true;
      }
      if ((match = pathname.match(/^\/api\/media\/items\/(\d+)\/render$/)) && method === "POST") {
        if (!needManage(res, requestUser)) return true;
        sendJson(res, 200, await queueRender(Number(match[1]), by));
        return true;
      }
      if ((match = pathname.match(/^\/api\/media\/items\/(\d+)$/)) && method === "PATCH") {
        if (!needManage(res, requestUser)) return true;
        const body = await readRequestJson(req);
        const fields = {};
        if (body.title !== undefined) fields.title = text(body.title, 200);
        if (body.caption_fr !== undefined) fields.caption_fr = text(body.caption_fr, 4000);
        if (body.caption_en !== undefined) fields.caption_en = text(body.caption_en, 4000);
        if (body.hashtags !== undefined) fields.hashtags = text(body.hashtags, 1000);
        if (body.archived !== undefined) fields.status = body.archived ? "archived" : "ready";
        const keys = Object.keys(fields);
        if (!keys.length) {
          sendJson(res, 400, { error: "Nothing to change" });
          return true;
        }
        const result = await dbQuery(
          `UPDATE media_items SET ${keys.map((key, index) => `${key} = $${index + 2}`).join(", ")}, updated_at = now() WHERE id = $1 RETURNING *`,
          [Number(match[1]), ...keys.map((key) => fields[key])],
        );
        sendJson(res, result.rows[0] ? 200 : 404, result.rows[0] ? { item: itemRow(result.rows[0]) } : { error: "Item not found" });
        return true;
      }

      // ---- Posts (by hand until the platform APIs are approved)
      if (pathname === "/api/media/posts" && method === "GET") {
        const values = [];
        const conditions = [];
        const from = text(url.searchParams.get("from"), 10);
        const to = text(url.searchParams.get("to"), 10);
        if (from) {
          values.push(from);
          conditions.push(`COALESCE(p.posted_at, p.scheduled_at) >= $${values.length}::date`);
        }
        if (to) {
          values.push(to);
          conditions.push(`COALESCE(p.posted_at, p.scheduled_at) < ($${values.length}::date + 1)`);
        }
        const result = await dbQuery(
          `SELECT p.*, i.title AS item_title, i.theme_key, i.recipe, i.outputs,
             COALESCE((SELECT jsonb_agg(to_jsonb(s.*) ORDER BY s.day) FROM social_post_stats s WHERE s.post_id = p.id), '[]'::jsonb) AS stats
           FROM social_posts p LEFT JOIN media_items i ON i.id = p.media_item_id
           ${conditions.length ? `WHERE ${conditions.join(" AND ")}` : ""}
           ORDER BY COALESCE(p.posted_at, p.scheduled_at) DESC NULLS LAST, p.id DESC LIMIT 500`,
          values,
        );
        sendJson(res, 200, { posts: result.rows.map(postRow) });
        return true;
      }
      if (pathname === "/api/media/posts" && method === "POST") {
        if (!needManage(res, requestUser)) return true;
        const body = await readRequestJson(req);
        const platforms = textList(body.platforms || [body.platform], 20, 6).filter((platform) => MEDIA_PLATFORMS.includes(platform));
        if (!platforms.length) {
          sendJson(res, 400, { error: "Choose a platform" });
          return true;
        }
        const posted = body.status === "published";
        const ids = [];
        for (const platform of platforms) {
          const result = await dbQuery(
            `INSERT INTO social_posts (media_item_id, platform, scheduled_at, posted_at, status, mode, post_url, caption, notes, created_by)
             VALUES ($1, $2, $3, $4, $5, 'manual', $6, $7, $8, $9) RETURNING id`,
            [Number(body.media_item_id) || null, platform, body.scheduled_at || null, posted ? body.posted_at || new Date().toISOString() : null,
              posted ? "published" : "planned", text(body.post_url, 500), text(body.caption, 4000), text(body.notes, 1000), by],
          );
          ids.push(Number(result.rows[0].id));
        }
        sendJson(res, 200, { ids });
        return true;
      }
      if ((match = pathname.match(/^\/api\/media\/posts\/(\d+)$/))) {
        const id = Number(match[1]);
        if (method === "PATCH") {
          if (!needManage(res, requestUser)) return true;
          const body = await readRequestJson(req);
          const sets = [];
          const values = [id];
          const add = (sql, value) => { values.push(value); sets.push(`${sql} = $${values.length}`); };
          if (body.scheduled_at !== undefined) add("scheduled_at", body.scheduled_at || null);
          if (body.caption !== undefined) add("caption", text(body.caption, 4000));
          if (body.notes !== undefined) add("notes", text(body.notes, 1000));
          if (body.post_url !== undefined) add("post_url", text(body.post_url, 500));
          if (body.status !== undefined && POST_STATUSES.includes(body.status)) {
            add("status", body.status);
            if (body.status === "approved") add("approved_by", by);
            if (body.status === "published") add("posted_at", body.posted_at || new Date().toISOString());
          }
          if (!sets.length) {
            sendJson(res, 400, { error: "Nothing to change" });
            return true;
          }
          const result = await dbQuery(`UPDATE social_posts SET ${sets.join(", ")}, updated_at = now() WHERE id = $1 RETURNING *`, values);
          sendJson(res, result.rows[0] ? 200 : 404, result.rows[0] ? { post: postRow(result.rows[0]) } : { error: "Post not found" });
          return true;
        }
        if (method === "DELETE") {
          if (!needManage(res, requestUser)) return true;
          // Only a post that never went out is removed.
          const result = await dbQuery("DELETE FROM social_posts WHERE id = $1 AND status IN ('planned', 'approved', 'cancelled') RETURNING id", [id]);
          sendJson(res, result.rows[0] ? 200 : 409, result.rows[0] ? { ok: true } : { error: "A published post stays (set it to cancelled if it was removed)" });
          return true;
        }
      }
      // Stats typed in by hand (from the platform's insights), per day.
      if ((match = pathname.match(/^\/api\/media\/posts\/(\d+)\/stats$/)) && method === "POST") {
        if (!needManage(res, requestUser)) return true;
        const body = await readRequestJson(req);
        const day = /^\d{4}-\d{2}-\d{2}$/.test(body.day || "") ? body.day : new Date().toISOString().slice(0, 10);
        const keys = ["views", "reach", "likes", "comments", "shares", "saves", "follows", "clicks"];
        await dbQuery(
          `INSERT INTO social_post_stats (post_id, day, ${keys.join(", ")}, source, updated_by)
           VALUES ($1, $2, ${keys.map((_, index) => `$${index + 3}`).join(", ")}, 'manual', $${keys.length + 3})
           ON CONFLICT (post_id, day) DO UPDATE SET ${keys.map((key) => `${key} = COALESCE(EXCLUDED.${key}, social_post_stats.${key})`).join(", ")}, source = 'manual', updated_by = EXCLUDED.updated_by, updated_at = now()`,
          [Number(match[1]), day, ...keys.map((key) => num(body[key])), by],
        );
        sendJson(res, 200, { ok: true });
        return true;
      }

      sendJson(res, 404, { error: "Unknown media route" });
      return true;
    } catch (error) {
      console.error("Media route failed:", pathname, error?.message || error);
      sendJson(res, error?.status || 500, { error: error?.message || "Media error" });
      return true;
    }
  };
}

// A media job claimed by a BEAST that then went away (crash, reboot, power)
// stays "claimed" forever: after 90 minutes (AI) / 30 minutes (others) it
// goes back to the queue, or fails when its attempts are used up.
export async function requeueStaleMediaJobs() {
  if (!isDatabaseEnabled()) return { requeued: 0, failed: 0 };
  const stale = `job_type LIKE 'media\\_%' AND status = 'claimed'
    AND claimed_at < now() - (CASE WHEN job_type LIKE 'media\\_ai%' THEN interval '90 minutes' ELSE interval '30 minutes' END)`;
  const requeued = await dbQuery(`UPDATE llm_jobs SET status = 'pending', agent_name = '', updated_at = now() WHERE ${stale} AND attempt_count < max_attempts RETURNING id`);
  const failed = await dbQuery(
    `UPDATE llm_jobs SET status = 'failed', error_text = 'The BEAST stopped answering during this job', finished_at = now(), updated_at = now()
     WHERE ${stale} AND attempt_count >= max_attempts RETURNING id, payload_json`,
  );
  for (const row of failed.rows) {
    const itemId = Number(row.payload_json?.item_id);
    if (itemId) await dbQuery("UPDATE media_items SET status = 'failed', error = 'The BEAST stopped answering during this job', updated_at = now() WHERE id = $1 AND job_id = $2", [itemId, row.id]);
  }
  return { requeued: requeued.rowCount, failed: failed.rowCount };
}

// HTTP for the BEAST (poller key, before the login check): the library scan
// results come in chunks here.
export function createMediaAgentRoutes({ sendJson, readRequestJson, checkAgentKey }) {
  return async function handleMediaAgentRoute(req, res, url) {
    if (url.pathname !== "/api/agent/media/assets" || req.method !== "POST") return false;
    if (!checkAgentKey(req, res)) return true;
    if (!isDatabaseEnabled()) {
      sendJson(res, 503, { error: "Database is not enabled" });
      return true;
    }
    try {
      sendJson(res, 200, await saveScannedAssets(await readRequestJson(req, 8 * 1024 * 1024)));
    } catch (error) {
      console.error("Media assets not saved:", error?.message || error);
      sendJson(res, 500, { error: error?.message || "Media assets error" });
    }
    return true;
  };
}
