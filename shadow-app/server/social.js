// Campagnes posting through the platforms' APIs -- for now Meta:
// Instagram (Business / Creator account) and the Facebook Page it is linked
// to. A Meta app in development mode may publish to the accounts of the
// people who have a role on that app, without App Review.
//
// Render (here) only does the light part: the Meta login (OAuth), keeping
// the page tokens, and queueing an approved post when it is due. The BEAST
// (media-poller-app) does the upload to Instagram / Facebook and fetches
// the stats; it asks for the token right before it needs it (agent route),
// so tokens never sit in the job queue.
import crypto from "node:crypto";
import { createLlmJob, dbQuery, isDatabaseEnabled } from "./db.js";

export const API_PLATFORMS = ["instagram", "facebook"];
const META_SCOPES = [
  "pages_show_list",
  "pages_read_engagement",
  "pages_manage_posts",
  "read_insights",
  "instagram_basic",
  "instagram_content_publish",
  "instagram_manage_insights",
  "business_management",
];

function text(value, max = 500) {
  return String(value ?? "").trim().slice(0, max);
}

// ---------------------------------------------------------------------------
// Meta app settings (camp_settings key "meta_app"; the secret never goes back
// to the browser)

async function readMetaApp() {
  if (!isDatabaseEnabled()) return {};
  const result = await dbQuery("SELECT value FROM camp_settings WHERE key = 'meta_app'");
  const value = result.rows[0]?.value || {};
  return {
    app_id: text(value.app_id, 40),
    app_secret: text(value.app_secret, 80),
    config_id: text(value.config_id, 40),
    graph_version: /^v\d+\.\d+$/.test(value.graph_version || "") ? value.graph_version : "v23.0",
  };
}

async function writeMetaApp(body, by) {
  const current = await readMetaApp();
  const next = {
    app_id: body.app_id !== undefined ? text(body.app_id, 40) : current.app_id,
    // An empty secret field keeps the stored one.
    app_secret: text(body.app_secret, 80) || current.app_secret,
    config_id: body.config_id !== undefined ? text(body.config_id, 40) : current.config_id,
    graph_version: /^v\d+\.\d+$/.test(body.graph_version || "") ? body.graph_version : current.graph_version || "v23.0",
  };
  await dbQuery(
    `INSERT INTO camp_settings (key, value, updated_by, updated_at) VALUES ('meta_app', $1, $2, now())
     ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_by = EXCLUDED.updated_by, updated_at = now()`,
    [JSON.stringify(next), by || ""],
  );
  return next;
}

function publicApp(app) {
  return { app_id: app.app_id, has_secret: Boolean(app.app_secret), config_id: app.config_id, graph_version: app.graph_version };
}

async function graphGet(app, path, params) {
  const url = new URL(`https://graph.facebook.com/${app.graph_version}/${path.replace(/^\//, "")}`);
  for (const [key, value] of Object.entries(params || {})) url.searchParams.set(key, value);
  const response = await fetch(url, { signal: AbortSignal.timeout(30000) });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok || payload.error) throw new Error(`Meta: ${payload.error?.message || response.status}`);
  return payload;
}

// ---------------------------------------------------------------------------
// Accounts

function accountRow(row) {
  return {
    id: Number(row.id),
    platform: row.platform,
    account_id: row.account_id,
    name: row.name,
    username: row.username,
    page_id: row.page_id,
    connected_by: row.connected_by,
    updated_at: row.updated_at,
  };
}

export async function listSocialAccounts() {
  if (!isDatabaseEnabled()) return [];
  const result = await dbQuery("SELECT * FROM social_accounts WHERE NOT disabled ORDER BY platform, name");
  return result.rows.map(accountRow);
}

// The Meta login gave a user token: keep a page token per Facebook Page and
// the Instagram account linked to each page (page tokens from a long-lived
// user token don't expire).
async function saveMetaAccounts(app, userToken, by) {
  const long = await graphGet(app, "oauth/access_token", {
    grant_type: "fb_exchange_token", client_id: app.app_id, client_secret: app.app_secret, fb_exchange_token: userToken,
  });
  const pages = await graphGet(app, "me/accounts", {
    access_token: long.access_token, limit: "100",
    fields: "id,name,access_token,instagram_business_account{id,username,name}",
  });
  const saved = [];
  for (const page of pages.data || []) {
    await upsertAccount({ platform: "facebook", account_id: page.id, name: page.name, username: "", token: page.access_token, page_id: page.id }, by);
    saved.push(`Facebook: ${page.name}`);
    const ig = page.instagram_business_account;
    if (ig?.id) {
      await upsertAccount({ platform: "instagram", account_id: ig.id, name: ig.name || ig.username || "", username: ig.username || "", token: page.access_token, page_id: page.id }, by);
      saved.push(`Instagram: @${ig.username || ig.id}`);
    }
  }
  return saved;
}

async function upsertAccount(account, by) {
  await dbQuery(
    `INSERT INTO social_accounts (platform, account_id, name, username, token, page_id, connected_by, disabled, updated_at)
     VALUES ($1, $2, $3, $4, $5, $6, $7, false, now())
     ON CONFLICT (platform, account_id) DO UPDATE SET name = EXCLUDED.name, username = EXCLUDED.username, token = EXCLUDED.token,
       page_id = EXCLUDED.page_id, connected_by = EXCLUDED.connected_by, disabled = false, updated_at = now()`,
    [account.platform, account.account_id, text(account.name, 200), text(account.username, 100), account.token, account.page_id || "", by || ""],
  );
}

// ---------------------------------------------------------------------------
// Scheduler (every minute, on the online server only): approved automatic
// posts that are due go to the BEAST.

function pickOutputs(outputs, platform) {
  const videos = outputs.filter((output) => output.kind === "video");
  const photos = outputs.filter((output) => output.kind === "photo");
  if (videos.length) {
    // Reels: the vertical one; else the first video.
    const video = videos.find((output) => output.variant === "9x16") || videos[0];
    const cover = outputs.find((output) => output.kind === "cover" && output.variant === video.variant);
    return { kind: "video", files: [video], cover: cover || null };
  }
  if (photos.length) {
    // Feed photos: 4:5 (or 1:1, else what there is), all cards of it -> carousel.
    const variant = ["4x5", "1x1", "9x16"].find((entry) => photos.some((output) => output.variant === entry)) || photos[0].variant;
    const files = photos.filter((output) => output.variant === variant).sort((a, b) => (a.index || 0) - (b.index || 0)).slice(0, platform === "instagram" ? 10 : 30);
    return { kind: files.length > 1 ? "carousel" : "photo", files, cover: null };
  }
  return null;
}

export async function queueDueSocialPosts() {
  if (!isDatabaseEnabled()) return { queued: 0 };
  const due = await dbQuery(
    `SELECT p.*, i.outputs, i.status AS item_status, a.page_id, a.id AS account_row
     FROM social_posts p
     JOIN media_items i ON i.id = p.media_item_id
     LEFT JOIN social_accounts a ON a.platform = p.platform AND a.account_id = p.account_id AND NOT a.disabled
     WHERE p.status = 'approved' AND p.mode = 'api' AND p.scheduled_at <= now()
     ORDER BY p.scheduled_at LIMIT 20`,
  );
  let queued = 0;
  for (const post of due.rows) {
    const fail = (message) => dbQuery("UPDATE social_posts SET status = 'failed', error = $2, updated_at = now() WHERE id = $1", [post.id, message]);
    if (!post.account_row) {
      await fail("The account is not connected (Settings > Instagram & Facebook)");
      continue;
    }
    if (post.item_status !== "ready") {
      await fail("The photo / video is not ready");
      continue;
    }
    const media = pickOutputs(post.outputs || [], post.platform);
    if (!media) {
      await fail("Nothing to post in this item");
      continue;
    }
    // Claim the post first, so a second run never queues it twice.
    const claim = await dbQuery("UPDATE social_posts SET status = 'publishing', error = '', updated_at = now() WHERE id = $1 AND status = 'approved' RETURNING id", [post.id]);
    if (!claim.rows[0]) continue;
    const job = await createLlmJob({
      job_type: "media_publish",
      created_by: post.approved_by || post.created_by || "system",
      priority: 30,
      // Never tried twice by itself: a second try could post twice.
      max_attempts: 1,
      payload_json: {
        post_id: Number(post.id),
        platform: post.platform,
        account_id: post.account_id,
        page_id: post.page_id,
        caption: post.caption,
        media,
      },
    });
    await dbQuery("UPDATE social_posts SET publish_job_id = $2 WHERE id = $1", [post.id, job.id]);
    queued += 1;
  }
  return { queued };
}

// Twice a day: the BEAST fetches the stats of the posts it published in the
// last 35 days.
export async function queueSocialStatsJob() {
  if (!isDatabaseEnabled()) return { queued: false };
  const pending = await dbQuery("SELECT 1 FROM llm_jobs WHERE job_type = 'media_stats' AND status IN ('pending', 'claimed') LIMIT 1");
  if (pending.rows[0]) return { queued: false };
  const posts = await dbQuery(
    `SELECT p.id, p.platform, p.account_id, p.platform_post_id, p.posted_at, i.recipe
     FROM social_posts p LEFT JOIN media_items i ON i.id = p.media_item_id
     WHERE p.mode = 'api' AND p.status = 'published' AND p.platform_post_id <> '' AND p.posted_at > now() - interval '35 days'
     LIMIT 300`,
  );
  if (!posts.rows.length) return { queued: false };
  await createLlmJob({
    job_type: "media_stats",
    created_by: "system",
    priority: 5,
    max_attempts: 1,
    payload_json: { posts: posts.rows.map((row) => ({ post_id: Number(row.id), platform: row.platform, account_id: row.account_id, platform_post_id: row.platform_post_id })) },
  });
  return { queued: true, posts: posts.rows.length };
}

// /result of media_publish / media_stats (called from media.js).
export async function handleSocialJobResult(job) {
  const result = job.result_json || {};
  if (job.job_type === "media_publish") {
    await dbQuery(
      `UPDATE social_posts SET status = 'published', posted_at = now(), post_url = $2, platform_post_id = $3, error = '', updated_at = now()
       WHERE id = $1`,
      [Number(job.payload_json?.post_id), text(result.post_url, 500), text(result.platform_post_id, 120)],
    );
  } else if (job.job_type === "media_stats") {
    const keys = ["views", "reach", "likes", "comments", "shares", "saves", "follows", "clicks"];
    const day = new Date().toISOString().slice(0, 10);
    for (const row of (Array.isArray(result.stats) ? result.stats : []).slice(0, 500)) {
      const postId = Number(row.post_id);
      if (!postId) continue;
      const values = keys.map((key) => (Number.isFinite(Number(row[key])) && row[key] !== null && row[key] !== "" ? Math.round(Number(row[key])) : null));
      await dbQuery(
        `INSERT INTO social_post_stats (post_id, day, ${keys.join(", ")}, source, updated_by)
         VALUES ($1, $2, ${keys.map((_, index) => `$${index + 3}`).join(", ")}, 'api', 'beast')
         ON CONFLICT (post_id, day) DO UPDATE SET ${keys.map((key) => `${key} = COALESCE(EXCLUDED.${key}, social_post_stats.${key})`).join(", ")},
           source = 'api', updated_at = now()`,
        [postId, day, ...values],
      );
    }
  }
}

export async function handleSocialJobFail(job, errorText) {
  if (job.job_type !== "media_publish") return;
  await dbQuery(
    "UPDATE social_posts SET status = 'failed', error = $2, updated_at = now() WHERE id = $1 AND status = 'publishing'",
    [Number(job.payload_json?.post_id), `${text(errorText || job.error_text, 1500)} -- check on the platform whether it went out before trying again`],
  );
}

// ---------------------------------------------------------------------------
// HTTP (logged-in users): /api/social/...

const oauthStates = new Map();

export function createSocialRoutes({ sendJson, readRequestJson, requirePermission, PERMISSIONS, publicBaseUrl }) {
  const redirectUri = (req) => `${publicBaseUrl(req)}/api/social/meta/callback`;

  return async function handleSocialRoute(req, res, url, requestUser) {
    const pathname = url.pathname;
    if (!pathname.startsWith("/api/social/")) return false;
    if (!requirePermission(res, requestUser, PERMISSIONS.CAMPAIGN_VIEW)) return true;
    if (!isDatabaseEnabled()) {
      sendJson(res, 503, { error: "Posting needs the database" });
      return true;
    }
    let match;
    try {
      if (pathname === "/api/social/status" && req.method === "GET") {
        sendJson(res, 200, { meta_app: publicApp(await readMetaApp()), accounts: await listSocialAccounts(), redirect_uri: redirectUri(req) });
        return true;
      }
      if (pathname === "/api/social/meta/app" && req.method === "PUT") {
        if (!requirePermission(res, requestUser, PERMISSIONS.CAMPAIGN_MANAGE)) return true;
        sendJson(res, 200, { meta_app: publicApp(await writeMetaApp(await readRequestJson(req), requestUser.username)) });
        return true;
      }
      if (pathname === "/api/social/meta/auth-url" && req.method === "GET") {
        if (!requirePermission(res, requestUser, PERMISSIONS.CAMPAIGN_MANAGE)) return true;
        const app = await readMetaApp();
        if (!app.app_id || !app.app_secret) {
          sendJson(res, 400, { error: "Fill in the Meta app id and secret first" });
          return true;
        }
        const state = crypto.randomBytes(16).toString("hex");
        oauthStates.set(state, { user: requestUser.username, at: Date.now() });
        for (const [key, value] of oauthStates) if (Date.now() - value.at > 30 * 60 * 1000) oauthStates.delete(key);
        const login = new URL(`https://www.facebook.com/${app.graph_version}/dialog/oauth`);
        login.searchParams.set("client_id", app.app_id);
        login.searchParams.set("redirect_uri", redirectUri(req));
        login.searchParams.set("state", state);
        login.searchParams.set("response_type", "code");
        // "Facebook Login for Business" apps use a configuration id instead
        // of a scope list.
        if (app.config_id) login.searchParams.set("config_id", app.config_id);
        else login.searchParams.set("scope", META_SCOPES.join(","));
        sendJson(res, 200, { url: login.toString() });
        return true;
      }
      if (pathname === "/api/social/meta/callback" && req.method === "GET") {
        const page = (title, body) => {
          res.writeHead(200, { "content-type": "text/html; charset=utf-8" });
          res.end(`<!doctype html><meta charset="utf-8"><title>${title}</title><body style="font-family:sans-serif;padding:24px"><h2>${title}</h2>${body}<p>You can close this tab and go back to SnappySjaak (Campagnes &gt; Settings).</p></body>`);
        };
        const state = url.searchParams.get("state") || "";
        if (!oauthStates.has(state)) {
          page("Meta login expired", "<p>Start again with Connect in Campagnes &gt; Settings.</p>");
          return true;
        }
        oauthStates.delete(state);
        if (url.searchParams.get("error")) {
          page("Meta login cancelled", `<p>${text(url.searchParams.get("error_description") || url.searchParams.get("error"), 300).replace(/[<>&]/g, "")}</p>`);
          return true;
        }
        const app = await readMetaApp();
        const token = await graphGet(app, "oauth/access_token", {
          client_id: app.app_id, client_secret: app.app_secret, redirect_uri: redirectUri(req), code: url.searchParams.get("code") || "",
        });
        const saved = await saveMetaAccounts(app, token.access_token, requestUser.username);
        page(saved.length ? "Connected" : "No pages found",
          saved.length
            ? `<ul>${saved.map((line) => `<li>${line.replace(/[<>&]/g, "")}</li>`).join("")}</ul>`
            : "<p>The login did not give access to a Facebook Page. Log in again and tick the company's Page (and its Instagram account).</p>");
        return true;
      }
      if ((match = pathname.match(/^\/api\/social\/accounts\/(\d+)$/)) && req.method === "DELETE") {
        if (!requirePermission(res, requestUser, PERMISSIONS.CAMPAIGN_MANAGE)) return true;
        await dbQuery("UPDATE social_accounts SET disabled = true, token = '', updated_at = now() WHERE id = $1", [Number(match[1])]);
        sendJson(res, 200, { ok: true });
        return true;
      }
      sendJson(res, 404, { error: "Unknown social route" });
      return true;
    } catch (error) {
      console.error("Social route failed:", pathname, error?.message || error);
      if (pathname === "/api/social/meta/callback") {
        res.writeHead(500, { "content-type": "text/html; charset=utf-8" });
        res.end(`<p>Meta login failed: ${String(error?.message || error).replace(/[<>&]/g, "")}</p>`);
        return true;
      }
      sendJson(res, error?.status || 500, { error: error?.message || "Social error" });
      return true;
    }
  };
}

// Agent route (poller key): the BEAST asks for an account's token right
// before it publishes / reads stats.
export function createSocialAgentRoutes({ sendJson, checkAgentKey }) {
  return async function handleSocialAgentRoute(req, res, url) {
    if (url.pathname !== "/api/agent/social/token" || req.method !== "GET") return false;
    if (!checkAgentKey(req, res)) return true;
    const platform = text(url.searchParams.get("platform"), 20);
    const accountId = text(url.searchParams.get("account_id"), 60);
    const result = await dbQuery("SELECT * FROM social_accounts WHERE platform = $1 AND account_id = $2 AND NOT disabled", [platform, accountId]);
    const account = result.rows[0];
    if (!account?.token) {
      sendJson(res, 404, { error: "Account not connected" });
      return true;
    }
    const app = await readMetaApp();
    sendJson(res, 200, { platform, account_id: account.account_id, page_id: account.page_id, access_token: account.token, graph_version: app.graph_version });
    return true;
  };
}
